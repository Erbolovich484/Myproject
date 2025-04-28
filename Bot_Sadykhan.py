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
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.default import DefaultBotProperties

from aiohttp import web

# === Настройка логирования ===
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s'
)

# === Загрузка окружения ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")      # https://<app>.onrender.com/webhook
PORT           = int(os.getenv("PORT", "8000"))

# === FSM-состояния ===
class Form(StatesGroup):
    name       = State()
    pharmacy   = State()
    rating     = State()
    conclusion = State()

# === Загрузка критериев ===
_raw = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = _raw[_raw.iloc[:,0] == "Блок"].index[0] + 1
_df = _raw.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])
criteria = []
_last = None
for _, row in _df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else _last
    _last = block
    maxv = int(row["Макс"]) if pd.notna(row["Макс"]) and str(row["Макс"]).isdigit() else 10
    criteria.append({
        "block": block,
        "criterion": row["Критерий"],
        "requirement": row["Требование"],
        "max": maxv
    })

# === Утилиты ===
def now_ts() -> str:
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy: str, name: str, ts: str, score: int, max_score: int):
    first = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if first:
            w.writerow(["Дата","Аптека","ФИО","Баллы","Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# === Инициализация бота ===
session = AiohttpSession()
bot = Bot(
    token=API_TOKEN,
    session=session,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML)
)
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)

# === /start ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    logging.debug(f"CMD /start from {msg.from_user.id}")
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "© ИТ-служба «Садыхан». Интеллектуальная собственность.\n"
        "Этот бот поможет пройти чек-лист и отправит итоговый отчёт в QA-чат.\n\n"
        "🏁 Введите ваше ФИО:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)

# === Сбор ФИО ===
@dp.message(Form.name)
async def name_handler(msg: types.Message, state: FSMContext):
    logging.debug(f"Received name: {msg.text!r}")
    await state.update_data(name=msg.text.strip(), step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# === Сбор аптеки ===
@dp.message(Form.pharmacy)
async def pharm_handler(msg: types.Message, state: FSMContext):
    logging.debug(f"Received pharmacy: {msg.text!r}")
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# === Callback для inline-кнопок ===
@dp.callback_query()
async def cb_handler(cb: types.CallbackQuery, state: FSMContext):
    logging.debug(f"Callback received: {cb.data}")
    await cb.answer()  # сброс таймаута
    data = await state.get_data()
    step = data.get("step", 0)

    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        if step < len(criteria):
            data.setdefault("data", []).append({"crit": criteria[step], "score": score})
            data["step"] += 1
            await state.set_data(data)
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

# === Отправка вопросов (с логами и защитой) ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)
    logging.debug(f"send_question: step={step}/{total} to {chat_id}")

    if step >= total:
        await bot.send_message(chat_id, "✅ Проверка завершена. Введите, пожалуйста, вывод по аптеке:")
        return await state.set_state(Form.conclusion)

    c = criteria[step]
    logging.debug(f"Criterion #{step+1}: block={c['block']!r}, max={c['max']}")

    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )

    kb = InlineKeyboardBuilder()
    start = 0 if c["max"] == 1 else 1
    for i in range(start, c["max"]+1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    try:
        await bot.send_message(
            chat_id,
            text,
            parse_mode=ParseMode.HTML,
            reply_markup=kb.as_markup()
        )
        logging.debug("Question sent successfully")
    except Exception as e:
        logging.error(f"Failed to send question #{step+1}: {e}", exc_info=True)

# === Сбор вывода и генерация отчёта ===
@dp.message(Form.conclusion)
async def conclusion_handler(msg: types.Message, state: FSMContext):
    logging.debug(f"Received conclusion: {msg.text!r}")
    await state.update_data(conclusion=msg.text.strip())
    data = await state.get_data()

    ts      = data["start"]
    name    = data["name"]
    pharm   = data["pharmacy"]
    answers = data["data"]
    concl   = data.get("conclusion", "")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    title = (
        f"Отчёт по проверке аптеки\n"
        f"Исполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    headers = ["Блок","Критерий","Требование","Баллы","Макс","Дата проверки"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    row = 6
    total_score = 0
    total_max   = 0
    for it in answers:
        c = it["crit"]
        s = it["score"]
        ws.cell(row,1,c["block"])
        ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"])
        ws.cell(row,4,s)
        ws.cell(row,5,c["max"])
        ws.cell(row,6,ts)
        total_score += s
        total_max   += c["max"]
        row += 1

    ws.cell(row+1, 3, "ИТОГО:")
    ws.cell(row+1, 4, total_score)
    ws.cell(row+2, 3, "Максимум:")
    ws.cell(row+2, 4, total_max)

    ws.cell(row+4, 1, "Вывод аудитора:")
    ws.merge_cells(start_row=row+4, start_column=2, end_row=row+4, end_column=7)
    ws.cell(row+4, 2, concl)

    fn = f"{pharm}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d%m%Y')}.xlsx".replace(" ","_")
    wb.save(fn)

    with open(fn, "rb") as f:
        await bot.send_document(QA_CHAT_ID, FSInputFile(f, filename=fn))
    with open(fn, "rb") as f:
        await bot.send_document(msg.chat.id, FSInputFile(f, filename=fn))

    os.remove(fn)
    log_csv(pharm, name, ts, total_score, total_max)

    await msg.answer("🎉 Отчёт отправлен в QA-чат и вам.\nЧтобы пройти снова — /start")
    await state.clear()

# === Webhook & healthcheck ===
async def handle_webhook(request: web.Request):
    data = await request.json()
    upd  = Update(**data)
    await dp.feed_update(bot=bot, update=upd)
    return web.Response(text="OK")

async def health(request: web.Request):
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    logging.debug("Setting webhook")
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True)

async def on_cleanup(app: web.Application):
    logging.debug("Deleting webhook & closing storage")
    await bot.delete_webhook()
    await storage.close()

app = web.Application()
app.router.add_get("/", health)
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__ == "__main__":
    logging.info("Starting app")
    web.run_app(app, host="0.0.0.0", port=PORT)
