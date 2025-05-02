import logging
import os
import csv
import pytz
from datetime import datetime

from dotenv import load_dotenv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.redis import RedisStorage
from aiogram.types import FSInputFile, Update
from aiogram.client.session.aiohttp import AiohttpSession

from aiohttp import web

# ——— Конфигурация ———
load_dotenv()
API_TOKEN     = os.getenv("API_TOKEN")           # ваш токен
QA_CHAT_ID    = int(os.getenv("QA_CHAT_ID"))     # куда дублировать отчёт
BASE_URL      = os.getenv("BASE_URL")            # https://ваше-приложение.onrailway.app
WEBHOOK_PATH  = f"/webhook/{API_TOKEN}"
WEBHOOK_URL   = BASE_URL + WEBHOOK_PATH
PORT          = int(os.getenv("PORT", 8080))
REDIS_URL     = os.getenv("REDIS_URL")           # из Railway Add‑ons > Redis

if not all([API_TOKEN, QA_CHAT_ID, BASE_URL, REDIS_URL]):
    raise RuntimeError("Не все обязательные переменные окружения установлены!")

# ——— FSM ———
class Form(StatesGroup):
    name       = State()
    pharmacy   = State()
    rating     = State()
    comment    = State()

# ——— Загрузка чек-листа ———
df = pd.read_excel(os.getenv("CHECKLIST_PATH", "template.xlsx"), sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])
criteria = []
last_block = None
for _, r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс"]) if pd.notna(r["Макс"]) and str(r["Макс"]).isdigit() else 10
    criteria.append({"block": blk, "criterion": r["Критерий"], "requirement": r["Требование"], "max": maxv})

# ——— Утилиты ———
def now_ts():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(ph, nm, ts, sc, mx):
    path = os.getenv("LOG_PATH", "log.csv")
    exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","ФИО","Баллы","Макс"])
        w.writerow([ts, ph, nm, sc, mx])

# ——— Бот и диспетчер ———
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

storage = RedisStorage.from_url(REDIS_URL)
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=storage)

# ——— Хэндлеры ———
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    logger.info(f"CMD /start from {msg.from_user.id}")
    await state.clear()
    await msg.answer(
        "📋 <b>Чек‑лист аптек</b>\n"
        "Пожалуйста, отвечайте вдумчиво. Это интеллектуальная собственность ИТ‑Садыхан.\n\n"
        "🏁 Введите ваше ФИО:",
    )
    await state.set_state(Form.name)

@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    name = msg.text.strip()
    ts = now_ts()
    await state.update_data(name=name, step=0, answers=[], start=ts)
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    pharm = msg.text.strip()
    await state.update_data(pharmacy=pharm)
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    logger.info(f"Callback: {cb.data} from {cb.from_user.id}")
    await cb.answer()  # чтобы Telegram не жаловался
    data = await state.get_data()
    step = data["step"]

    # Ввод оценки
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        data["answers"].append({"crit": criteria[step], "score": score})
        data["step"] += 1
        await state.set_data(data)
        # отредактировать сообщение оценки
        await bot.edit_message_text(cb.from_user.id, cb.message.message_id,
                                    f"✅ Оценка: {score} {'⭐'*score}")
        return await send_question(cb.from_user.id, state)

@dp.message(Form.rating)
async def timeout_handler(msg: types.Message, state: FSMContext):
    # если пользователь пишет текст, а не нажимает inline-кнопку — считаем как «пропуск»
    data = await state.get_data()
    score = 0
    step = data["step"]
    data["answers"].append({"crit": criteria[step], "score": score})
    data["step"] += 1
    await state.set_data(data)
    await send_question(msg.chat.id, state)

async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step, total = data["step"], len(criteria)
    if step >= total:
        # переходим к комментарию
        await state.set_state(Form.comment)
        return await bot.send_message(chat_id, "📝 Напишите, пожалуйста, Ваш итоговый вывод по аптеке:")
    crit = criteria[step]
    kb = InlineKeyboardBuilder()
    start = 0 if crit["max"]==1 else 1
    for i in range(start, crit["max"]+1):
        kb.button(str(i), callback_data=f"score_{i}")
    kb.adjust(5)
    text = (f"<b>Вопрос {step+1}/{total}</b>\n"
            f"<b>Блок:</b> {crit['block']}\n"
            f"<b>Критерий:</b> {crit['criterion']}\n"
            f"<b>Требование:</b> {crit['requirement']}\n"
            f"<b>Макс:</b> {crit['max']}")
    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    data = await state.get_data()
    data["comment"] = msg.text.strip()
    await state.set_data(data)
    await bot.send_message(msg.chat.id, "Генерирую отчёт…")
    await make_report(msg.chat.id, data)
    await state.clear()

async def make_report(user_id: int, data: dict):
    wb = load_workbook(os.getenv("TEMPLATE_PATH","template.xlsx"))
    ws = wb.active
    # заголовок
    ws.merge_cells("A1:G2")
    ws["A1"] = (f"Отчёт по проверке аптек\n"
                f"Исполнитель: {data['name']}\n"
                f"Дата: {datetime.strptime(data['start'],'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M')}")
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = data["pharmacy"]
    # шапка
    headers = ["Блок","Критерий","Требование","Оценка","Макс","Комментарий","Вывод"]
    for idx, h in enumerate(headers, start=1):
        ws.cell(row=5, column=idx, value=h).font = Font(bold=True)
    # строки
    row = 6; total=0; total_max=0
    for ans in data["answers"]:
        crit, sc = ans["crit"], ans["score"]
        ws.cell(row,1,crit["block"])
        ws.cell(row,2,crit["criterion"])
        ws.cell(row,3,crit["requirement"])
        ws.cell(row,4,sc)
        ws.cell(row,5,crit["max"])
        ws.cell(row,6,"—")
        ws.cell(row,7,data["comment"] if row==6 else "")
        total += sc; total_max+=crit["max"]
        row +=1
    ws.cell(row,3,"ИТОГО:"); ws.cell(row,4,total)
    ws.cell(row+1,3,"Максимум:"); ws.cell(row+1,4,total_max)
    # сохранить
    fname = f"{data['pharmacy']}_{data['name']}_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx".replace(" ","_")
    wb.save(fname)

    # отправить самому проверяющему
    with open(fname,"rb") as f:
        await bot.send_document(user_id, FSInputFile(f, filename=fname))
    # и в QA-чат
    with open(fname,"rb") as f:
        await bot.send_document(QA_CHAT_ID, FSInputFile(f, filename=fname))

    os.remove(fname)
    log_csv(data["pharmacy"], data["name"], data["start"], total, total_max)
    await bot.send_message(user_id, "✅ Отчёт готов! Чтобы начать заново — /start")

# ——— Webhook & запуск ———
app = web.Application()
app.router.add_post(WEBHOOK_PATH, lambda r: handle_update(r))
app.router.add_get("/", lambda r: web.Response(text="OK"))
async def handle_update(request: web.Request):
    data = await request.json()
    upd  = Update(**data)
    await dp.feed_update(bot, upd)
    return web.Response(text="OK")

async def on_startup(app):
    logger.info("Setting webhook to %s", WEBHOOK_URL)
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True)

async def on_shutdown(app):
    await bot.delete_webhook()
    await storage.close()

app.on_startup.append(on_startup)
app.on_cleanup.append(on_shutdown)

if __name__ == "__main__":
    web.run_app(app, host="0.0.0.0", port=PORT)
