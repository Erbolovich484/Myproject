import logging, os, csv, pytz
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
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update
from aiohttp import web

# === Загрузка конфигов ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
PORT           = int(os.getenv("PORT", 8000))

# === FSM-состояния ===
class Form(StatesGroup):
    name    = State()
    pharmacy= State()
    rating  = State()
    comment = State()

# === Читаем критерии из Excel ===
df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0]+1
df = df.iloc[start_i:,:8].dropna(subset=[1,2]).reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]

criteria = []
last = None
for _,r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last
    last = blk
    maxv = int(r["Макс"]) if str(r["Макс"]).isdigit() else 10
    criteria.append({
        "block": blk,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# === Утилиты ===
def now_ts():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharm,name,ts,score,total):
    first = not os.path.exists(LOG_PATH)
    with open(LOG_PATH,"a",newline="",encoding="utf-8") as f:
        w = csv.writer(f)
        if first:
            w.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        w.writerow([ts,pharm,name,score,total])

# === Инициализация бота ===
logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=MemoryStorage())

# === Команда /start ===
@dp.message(F.text=="/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "<b>📋 Чек‑лист посещения аптек</b>\n\n"
        "Этот бот — интеллектуальная собственность ИТ «Садыхан».  \n"
        "Заполняйте чек‑лист вдумчиво и внимательно:  \n"
        "- inline‑кнопки для быстрой оценки;  \n"
        "- если оценка займёт больше минуты — после всех баллов вы сможете написать вывод ручкой.\n\n"
        "Введите ваше ФИО для авторизации:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)

# === /id для отладки ===
@dp.message(F.text=="/id")
async def cmd_id(msg: types.Message):
    await msg.answer(f"<code>{msg.chat.id}</code>")

# === Сброс FSM ===
@dp.message(F.text=="/сброс")
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново")

# === Обработка ФИО ===
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    name = msg.text.strip()
    await state.update_data(name=name, step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# === Обработка названия аптеки ===
@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# === Общий хэндлер callback_query ===
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    step = data.get("step",0)
    total= len(criteria)
    # если уже все оценили — просто acknowledge
    if step>=total:
        return await cb.answer()
    # парсим оценку
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_",1)[1])
        record = {"crit": criteria[step], "score": score}
        data["data"].append(record)
        data["step"] += 1
        await state.update_data(**data)
        # редактируем исходное сообщение
        await bot.edit_message_text(
            f"✅ Оценка: {score} {'⭐'*score}",
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id
        )
        # отправляем следующий
        return await send_question(cb.from_user.id, state)
    # навигация «Назад»
    if cb.data=="prev" and step>0:
        data["step"]-=1
        data["data"].pop()
        await state.update_data(**data)
        return await send_question(cb.from_user.id, state)

# === Функция отправки следующего вопроса или финального промпта ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total= len(criteria)

    # если всё — переходим к комментарию
    if step>=total:
        await bot.send_message(
            chat_id,
            "✅ Все оценки поставлены!\n\n"
            "📝 Теперь напишите, пожалуйста, ваши выводы по аптеке:",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(Form.comment)
        return

    c = criteria[step]
    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )
    kb = InlineKeyboardBuilder()
    start = 0 if c["max"]==1 else 1
    for i in range(start, c["max"]+1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# === Обработка текстового комментария ===
@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    data = await state.get_data()
    data["comment"] = msg.text.strip()
    await state.update_data(**data)
    await msg.answer("⌛ Формирую отчёт…")
    await make_report(msg.chat.id, data)
    await state.clear()

# === Генерация и отправка отчёта ===
async def make_report(chat_id: int, data):
    name     = data["name"]
    ts       = data["start"]
    pharmacy = data["pharmacy"]
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Заголовок
    title = (f"Отчёт по проверке аптеки\n"
             f"Исполнитель: {name}\n"
             f"Дата и время: {ts}")
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharmacy

    # Шапка таблицы
    headers = ["Блок","Критерий","Требование","Оценка","Макс","Коммент.","Дата проверки"]
    for idx,h in enumerate(headers, start=1):
        cell = ws.cell(5, idx, h)
        cell.font = Font(bold=True)

    # Заполняем строки
    row = 6
    total_score = 0
    total_max   = 0
    for rec in data["data"]:
        c = rec["crit"]
        sc= rec["score"]
        ws.cell(row,1, c["block"])
        ws.cell(row,2, c["criterion"])
        ws.cell(row,3, c["requirement"])
        ws.cell(row,4, sc)
        ws.cell(row,5, c["max"])
        ws.cell(row,6, "")  # можно сюда вставить индивидуальный прим.
        ws.cell(row,7, ts)
        total_score += sc
        total_max   += c["max"]
        row += 1

    # Итого
    ws.cell(row+1,3, "ИТОГО:")
    ws.cell(row+1,4, total_score)
    ws.cell(row+2,3, "Максимум:")
    ws.cell(row+2,4, total_max)

    # Ваш комментарий внизу
    ws.cell(row+4,1, "Вывод проверяющего:")
    ws.cell(row+5,1, data.get("comment",""))

    # Сохраняем файл
    fn = f"{pharmacy}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y_%H%M')}.xlsx"
    fn = fn.replace(" ","_")
    wb.save(fn)

    # 1) Отправляем пользователю
    with open(fn,"rb") as f:
        await bot.send_document(chat_id, FSInputFile(f, fn))
    # 2) И в QA‑чат
    with open(fn,"rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f, fn))

    # Логируем
    log_csv(pharmacy, name, ts, total_score, total_max)

    # Финальное сообщение
    await bot.send_message(chat_id,
        "✅ Отчёт сформирован и отправлен.\n"
        "Для новой проверки — /start"
    )
    os.remove(fn)

# === Webhook & запуск ===
async def handle_webhook(request: web.Request):
    update = Update(**await request.json())
    await dp.feed_update(bot, update)
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    # если нужен Webhook:
    # await bot.set_webhook(os.getenv("WEBHOOK_URL"))
    pass

app = web.Application()
app.router.add_post("/webhook", handle_webhook)
app.router.add_get("/", lambda r: web.Response(text="OK"))

if __name__=="__main__":
    logging.getLogger("aiohttp.access").setLevel(logging.WARNING)
    web.run_app(app, host="0.0.0.0", port=PORT)
