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

# === ЗАГРУЗКА ПЕРЕМЕННЫХ ОКРУЖЕНИЯ ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")      # https://<your-app>.onrender.com/webhook
PORT           = int(os.getenv("PORT", "8000"))

# === FSM СТЕЙТЫ ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === ЗАГРУЗКА КРИТЕРИЕВ ИЗ EXCEL ===
criteria_df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_i = criteria_df[criteria_df.iloc[:, 0] == "Блок"].index[0] + 1
criteria_df = criteria_df.iloc[start_i:, :8].reset_index(drop=True)
criteria_df.columns = [
    "Блок", "Критерий", "Требование", "Оценка",
    "Макс. значение", "Примечание", "Дата проверки", "Дата исправления"
]
criteria_df = criteria_df.dropna(subset=["Критерий", "Требование"])

criteria_list = []
last_block = None
for _, r in criteria_df.iterrows():
    block = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = block
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria_list.append({
        "block": block,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def get_astana_time():
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")

def log_submission(pharmacy, name, ts, score, max_score):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата", "Аптека", "ФИО", "Баллы", "Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# === ИНИЦИАЛИЗАЦИЯ БОТА И ДИСПЕТЧЕРА ===
session = AiohttpSession()
bot = Bot(
    token=API_TOKEN,
    session=session,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML)
)
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)

# === ХЕНДЛЕРЫ ЧАТА ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    intro = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Заполняйте внимательно, отчёт придёт автоматически.\n\n"
        "🏁 Начнём!"
    )
    await msg.answer(intro, parse_mode=ParseMode.HTML)
    await msg.answer("Введите ваше ФИО:")
    await state.set_state(Form.name)

@dp.message(Form.name)
async def name_handler(msg: types.Message, state: FSMContext):
    await state.update_data(name=msg.text.strip(), step=0, data=[], start=get_astana_time())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def pharmacy_handler(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_criterion(msg.chat.id, state)

@dp.callback_query(F.data.startswith("score_") | F.data == "prev")
async def score_handler(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()
    data = await state.get_data()
    step = data["step"]

    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.set_data(data)
        return await send_criterion(cb.from_user.id, state)

    score = int(cb.data.split("_")[1])
    if step < len(criteria_list):
        data.setdefault("data", []).append({"crit": criteria_list[step], "score": score})
        data["step"] += 1
        await state.set_data(data)

    await bot.edit_message_text(
        chat_id=cb.message.chat.id,
        message_id=cb.message.message_id,
        text=f"✅ Оценка: {score} {'⭐'*score}"
    )
    await send_criterion(cb.from_user.id, state)

async def send_criterion(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria_list)

    if step >= total:
        await bot.send_message(chat_id, "Проверка завершена. Формируем отчёт…")
        return await generate_and_send(chat_id, data)

    c = criteria_list[step]
    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )

    kb = InlineKeyboardBuilder()
    start = 0 if c["max"] == 1 else 1
    for i in range(start, c["max"] + 1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, reply_markup=kb.as_markup(), parse_mode=ParseMode.HTML)

async def generate_and_send(chat_id: int, data):
    name  = data["name"]
    ts    = data["start"]
    pharm = data.get("pharmacy", "Без названия")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    title = (
        f"Отчёт по проверке аптеки\nИсполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    headers = ["Блок","Критерий","Требование","Баллы","Макс","Примечание","Дата"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    row = 6
    total = 0
    total_max = 0
    for item in data["data"]:
        c = item["crit"]; sc = item["score"]
        ws.cell(row, 1, c["block"])
        ws.cell(row, 2, c["criterion"])
        ws.cell(row, 3, c["requirement"])
        ws.cell(row, 4, sc)
        ws.cell(row, 5, c["max"])
        ws.cell(row, 7, ts)
        total += sc
        total_max += c["max"]
        row += 1

    ws.cell(row+1, 3, "ИТОГО:");    ws.cell(row+1, 4, total)
    ws.cell(row+2, 3, "Максимум:"); ws.cell(row+2, 4, total_max)

    filename = f"{pharm}_{name}_{datetime.strptime(ts, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx".replace(" ", "_")
    wb.save(filename)
    with open(filename, "rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f))
    os.remove(filename)
    log_submission(pharm, name, ts, total, total_max)
    await bot.send_message(chat_id, "Готово! /start — чтобы заново.")

# === Webhook handler и Health-check===
async def handle_webhook(request: web.Request):
    data = await request.json()
    update = Update(**data)
    await dp.feed_update(update=update)
    return web.Response(text="OK")

async def health(request: web.Request):
    return web.Response(text="OK")

# === Запуск AioHTTP и Webhook ===
async def on_startup(app: web.Application):
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True)

async def on_cleanup(app: web.Application):
    await bot.delete_webhook()
    await storage.close()

app = web.Application()
app.router.add_get("/", health)
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    web.run_app(app, host="0.0.0.0", port=PORT)
