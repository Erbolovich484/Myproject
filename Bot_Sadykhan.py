import os
import csv
import pytz
import asyncio
import logging
import json
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

# === SENIOR NOTE: Загружаем .env и валидируем env vars ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = int(os.getenv("QA_CHAT_ID", "0"))
PUBLIC_DOMAIN  = os.getenv("RAILWAY_PUBLIC_DOMAIN")  # e.g. web-production-xxxx.up.railway.app
if not all([API_TOKEN, QA_CHAT_ID, PUBLIC_DOMAIN]):
    raise RuntimeError("Missing required env vars: API_TOKEN, QA_CHAT_ID or RAILWAY_PUBLIC_DOMAIN")

TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "checklist.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")

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

# === Load checklist ===
try:
    _df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
except Exception as e:
    logger.error(f"Failed to load checklist: {e}")
    raise

start_i = _df[_df.iloc[:,0] == "Блок"].index[0] + 1
_df = _df.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])

criteria = []
last_block = None
for _, r in _df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс"]) if pd.notna(r["Макс"]) and str(r["Макс"]).isdigit() else 10
    criteria.append({"block": blk, "criterion": r["Критерий"], "requirement": r["Требование"], "max": maxv})
TOTAL = len(criteria)

# === Utils ===
def now_str(fmt="%Y-%m-%d_%H-%M-%S"):
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime(fmt)

def log_csv(ph, nm, ts, sc, mx):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        w.writerow([ts, ph, nm, sc, mx])

# === Bot & Dispatcher ===
bot = Bot(token=API_TOKEN)
dp  = Dispatcher(storage=MemoryStorage())

# === Command Handlers ===
@dp.message(Command("start"))
async def cmd_start(msg: types.Message, state: FSMContext):
    logger.info("cmd_start triggered")
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
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id = <code>{msg.chat.id}</code>", parse_mode=ParseMode.HTML)

@dp.message(Command("лог"))
async def cmd_log(msg: types.Message):
    if os.path.exists(LOG_PATH):
        await msg.answer_document(InputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не создан.")

@dp.message(Command("сброс"))
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново.")

# === FSM Steps ===
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    await state.update_data(name=msg.text.strip(), step=0, answers=[], start=now_str("%Y-%m-%d %H:%M:%S"))
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
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
            logger.warning(f"Edit message failed: {e}")

        if data["step"] >= TOTAL:
            await bot.send_message(cb.from_user.id, "✅ Проверка завершена. Введите вывод проверяющего (или «—»):")
            return await state.set_state(Form.comment)

        return await send_question(cb.from_user.id, state)

@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    await state.update_data(comment=msg.text.strip())
    await msg.answer("Формирую отчёт…")
    data = await state.get_data()
    await make_report(msg.chat.id, data)
    await state.clear()

# === Question Sender ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    crit = criteria[step]

    kb = InlineKeyboardBuilder()
    start = 0 if crit["max"] == 1 else 1
    for i in range(start, crit["max"] + 1):
        kb.button(str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button("◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(
        chat_id,
        (f"<b>Вопрос {step+1} из {TOTAL}</b>\n\n"
         f"<b>Блок:</b> {crit['block']}\n"
         f"<b>Критерий:</b> {crit['criterion']}\n"
         f"<b>Требование:</b> {crit['requirement']}\n"
         f"<b>Макс. балл:</b> {crit['max']}"),
        reply_markup=kb.as_markup(),
        parse_mode=ParseMode.HTML
    )

# === Report Generator ===
async def make_report(user_chat: int, data: dict):
    name    = data["name"]
    pharm   = data["pharmacy"]
    ts      = data["start"]
    comment = data["comment"]
    answers = data["answers"]

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    title = f"Отчёт: {pharm} — {name} ({ts.split()[0]})"
    ws.merge_cells("A1:G2")
    ws["A1"] = title; ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    hdr = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки"]
    for idx, h in enumerate(hdr, 1):
        ws.cell(row=5, column=idx, value=h).font = Font(bold=True)

    row = 6; total = 0; max_total = 0
    for it in answers:
        c = it["crit"]; sc = it["score"]
        ws.cell(row, 1, c["block"])
        ws.cell(row, 2, c["criterion"])
        ws.cell(row, 3, c["requirement"])
        ws.cell(row, 4, sc)
        ws.cell(row, 5, c["max"])
        ws.cell(row, 7, ts)
        total += sc; max_total += c["max"]
        row += 1

    ws.cell(row+1, 3, "ИТОГО:");    ws.cell(row+1, 4, total)
    ws.cell(row+2, 3, "Максимум:"); ws.cell(row+2, 4, max_total)
    ws.cell(row+4, 1, "Вывод проверяющего:"); ws.cell(row+4, 2, comment)

    fn = f"{pharm}_{name}_{now_str()}.xlsx".replace(" ", "_")
    wb.save(fn)

    for chat in (user_chat, QA_CHAT_ID):
        try:
            await bot.send_document(chat, InputFile(fn))
        except Exception as e:
            logger.error(f"Failed to send report to {chat}: {e}")
    os.remove(fn)

    log_csv(pharm, name, ts, total, max_total)

# === Webhook Setup & Application ===
async def on_startup(app: web.Application):
    await bot.delete_webhook(drop_pending_updates=True)
    await bot.set_webhook(WEBHOOK_URL)
    logger.info(f"Webhook установлен: {WEBHOOK_URL}")

async def on_shutdown(app: web.Application):
    await bot.delete_webhook()
    logger.info("Webhook снят.")

async def handle_update(request: web.Request) -> web.Response:
    raw = await request.text()
    logger.info(f"Incoming raw update: {raw[:500]}")
    try:
        data = json.loads(raw)
        update = types.Update(**data)
        await dp.feed_update(update)
    except Exception:
        logger.exception("Error processing update")
    return web.Response(status=200)

def build_app() -> web.Application:
    app = web.Application()
    app.router.add_post(WEBHOOK_PATH, handle_update)
    app.on_startup.append(on_startup)
    app.on_shutdown.append(on_shutdown)
    return app

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8080))
    web_app = build_app()
    web.run_app(web_app, host="0.0.0.0", port=port)
