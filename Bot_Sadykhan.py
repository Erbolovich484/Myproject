import os
import csv
import pytz
import asyncio
import logging
import json
import tempfile
from datetime import datetime
from dotenv import load_dotenv

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiohttp import web
from aiogram import Bot, Dispatcher, types
from aiogram.types import InputFile
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

# === Load environment variables ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = os.getenv("QA_CHAT_ID")
PUBLIC_DOMAIN  = os.getenv("RAILWAY_PUBLIC_DOMAIN")
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "checklist.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")

# Validate required environment variables
missing = []
for var in ("API_TOKEN","QA_CHAT_ID","PUBLIC_DOMAIN"):  
    if not locals().get(var):  
        missing.append(var)
if missing:
    raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

API_TOKEN  = API_TOKEN.strip()
QA_CHAT_ID = int(QA_CHAT_ID)

WEBHOOK_PATH = f"/webhook/{API_TOKEN}"
WEBHOOK_URL  = f"https://{PUBLIC_DOMAIN}{WEBHOOK_PATH}"

# === Logging ===
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === FSM States ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()
    comment  = State()

# === Load and parse checklist ===
try:
    df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
except Exception as e:
    logger.error(f"Failed loading checklist: {e}")
    raise

start_idx = df[df.iloc[:,0] == "Блок"].index[0] + 1
df = df.iloc[start_idx:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])

criteria = []
last_block = None
for _, row in df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else last_block
    last_block = block
    max_val = int(row["Макс"]) if pd.notna(row["Макс"]) and str(row["Макс"]).isdigit() else 10
    criteria.append({
        "block": block,
        "criterion": row["Критерий"],
        "requirement": row["Требование"],
        "max": max_val
    })
TOTAL = len(criteria)

# === Utility functions ===
def now_str(fmt: str = "%Y-%m-%d_%H-%M-%S") -> str:
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime(fmt)

def log_csv(pharmacy: str, name: str, timestamp: str, score: int, max_score: int) -> None:
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        writer.writerow([timestamp, pharmacy, name, score, max_score])

# === Bot & Dispatcher ===
bot = Bot(token=API_TOKEN)
dp  = Dispatcher(storage=MemoryStorage())

# === Command Handlers ===
@dp.message(Command("start"))
async def cmd_start(msg: types.Message, state: FSMContext) -> None:
    logger.info("/start received from %s", msg.from_user.id)
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Интеллектуальная собственность ИТ-Садыхан.\n"
        "Заполняйте вдумчиво — отчёт придёт вам и в QA-чат.\n\n"
        "🏁 Введите ваше ФИО:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)

@dp.message(Command("id"))
async def cmd_id(msg: types.Message) -> None:
    await msg.answer(f"Ваш chat_id = <code>{msg.chat.id}</code>", parse_mode=ParseMode.HTML)

@dp.message(Command("лог"))
async def cmd_log(msg: types.Message) -> None:
    if os.path.exists(LOG_PATH):
        await msg.answer_document(InputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не создан.")

@dp.message(Command("сброс"))
async def cmd_reset(msg: types.Message, state: FSMContext) -> None:
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново.")

# === FSM Step Handlers ===
semaphore = asyncio.Semaphore(1)

@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext) -> None:
    await state.update_data(name=msg.text.strip(), step=0, answers=[], start=now_str())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext) -> None:
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext) -> None:
    await cb.answer()
    data = await state.get_data()
    step = data["step"]

    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["answers"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        crit  = criteria[step]
        data["answers"].append({"crit": crit, "score": score})
        data["step"] += 1
        await state.set_data(data)
        try:
            await bot.edit_message_text(
                f"✅ Оценка: {score} {'⭐'*score}",
                cb.message.chat.id,
                cb.message.message_id
            )
        except Exception as e:
            logger.warning("Edit message failed: %s", e)

        if data["step"] >= TOTAL:
            await bot.send_message(cb.from_user.id, "✅ Проверка завершена. Введите вывод проверяющего (или «—»):")
            return await state.set_state(Form.comment)

        return await send_question(cb.from_user.id, state)

@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext) -> None:
    await state.update_data(comment=msg.text.strip())
    await msg.answer("Формирую отчёт…")
    data = await state.get_data()
    await make_report(msg.chat.id, data)
    await state.clear()

# === Question Sender ===
async def send_question(chat_id: int, state: FSMContext) -> None:
    try:
        data = await state.get_data()
        step = data["step"]
        crit = criteria[step]
        logger.info(f"send_question: step={step}, criterion={crit['criterion']}")

        kb = InlineKeyboardBuilder()
        start = 0 if crit["max"] == 1 else 1
        for i in range(start, crit["max"] + 1):
            kb.button(text=str(i), callback_data=f"score_{i}")
        if step > 0:
            kb.button(text="◀️ Назад", callback_data="prev")
        kb.adjust(5)

        async with semaphore:
            text = (
                f"<b>Вопрос {step+1} из {TOTAL}</b>\n\n"
                f"<b>Блок:</b> {crit['block']}\n"
                f"<b>Критерий:</b> {crit['criterion']}\n"
                f"<b>Требование:</b> {crit['requirement']}\n"
                f"<b>Макс. балл:</b> {crit['max']}"
            )
            await bot.send_message(
                chat_id,
                text,
                reply_markup=kb.as_markup(),
                parse_mode=ParseMode.HTML
            )
            await asyncio.sleep(0.1)
    except Exception as e:
        logger.exception("Error in send_question: %s", e)

# === Report Generator ===
async def make_report(user_chat: int, data: dict) -> None:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    title = f"Отчёт: {data['pharmacy']} — {data['name']} ({data['start'].split()[0]})"
    ws.merge_cells("A1:G2")
    ws["A1"] = title; ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = data['pharmacy']

    hdr = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки"]
    for idx, h in enumerate(hdr, 1):
        ws.cell(row=5, column=idx, value=h).font = Font(bold=True)

    row = 6
    total = 0
    max_total = 0
    for it in data['answers']:
        c = it['crit']
        sc = it['score']
        ws.cell(row, 1, c['block'])
        ws.cell(row, 2, c['criterion'])
        ws.cell(row, 3, c['requirement'])
        ws.cell(row, 4, sc)
        ws.cell(row, 5, c['max'])
        ws.cell(row, 7, data['start'])
        total += sc
        max_total += c['max']
        row += 1

    ws.cell(row+1, 3, "ИТОГО:"); ws.cell(row+1, 4, total)
    ws.cell(row+2, 3, "Максимум:"); ws.cell(row+2, 4, max_total)
    ws.cell(row+4, 1, "Вывод проверяющего:"); ws.cell(row+4, 2, data['comment'])

    # Save to a temporary file to avoid collisions
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    for chat in (user_chat, QA_CHAT_ID):
        try:
            # Send as InputFile with filename
            with open(tmp_path, 'rb') as f:
                input_file = InputFile(f, filename=os.path.basename(tmp_path))
                await bot.send_document(chat, input_file)
        except Exception as e:
            logger.error("Failed to send report to %s: %s", chat, e)
    os.remove(tmp_path)
    log_csv(data['pharmacy'], data['name'], data['start'], total, max_total)

# === Webhook Setup ===
async def handle_update(request: web.Request) -> web.Response:
    raw = await request.text()
    logger.info("Incoming raw update: %s", raw[:500])
    try:
        data = json.loads(raw)
        update = types.Update(**data)
        await dp.feed_update(bot, update)
    except Exception:
        logger.exception("Error processing update")
    return web.Response(status=200)

app = web.Application()
app.router.add_post(WEBHOOK_PATH, handle_update)
app.on_startup.append(lambda app: asyncio.create_task(bot.set_webhook(WEBHOOK_URL)))
app.on_shutdown.append(lambda app: asyncio.create_task(bot.delete_webhook()))

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8080))
    web.run_app(app, host="0.0.0.0", port=port)
