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

# === Настройки ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")
PORT           = int(os.getenv("PORT", "8000"))

ALLOWED_USERS = [
    "Николай Крылов", "Таждин Усейн", "Жанар Бөлтірік",
    "Шара Абдиева", "Тохтар Чарабасов", "*"
]

class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Читаем чек-лист ===
df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])

criteria = []
last = None
for _, r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last
    last = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({"block": blk, "criterion":r["Критерий"], "requirement":r["Требование"], "max":maxv})

def get_time():
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(ph, nm, ts, sc, mx):
    ex = os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if not ex:
            w.writerow(["Дата","Аптека","ФИО проверяющего","Баллы","Макс"])
        w.writerow([ts, ph, nm, sc, mx])

# === Инициализация ===
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)

# --- Общие команды ---
@dp.message(F.text == "/start")
async def cmd_start(m: types.Message, s: FSMContext):
    await s.clear()
    text = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Заполняйте внимательно, отчёт придёт автоматически.\n\n"
        "✅ По завершении — отчёт в Excel.\n"
        "🏁 Введите ваше ФИО:"
    )
    await m.answer(text)
    await s.set_state(Form.name)

@dp.message(F.text == "/id")
async def cmd_id(m: types.Message):
    await m.answer(f"Ваш chat_id: <code>{m.chat.id}</code>", parse_mode=ParseMode.HTML)

@dp.message(F.text == "/лог")
async def cmd_log(m: types.Message):
    if os.path.exists(LOG_PATH):
        await m.answer_document(FSInputFile(LOG_PATH))
    else:
        await m.answer("Лог ещё не сформирован.")

@dp.message(F.text == "/сброс")
async def cmd_reset(m: types.Message, s: FSMContext):
    await s.clear()
    await m.answer("Состояние сброшено. /start — чтобы начать заново.")

# --- Авторизация ---
@dp.message(Form.name)
async def proc_name(m: types.Message, s: FSMContext):
    u = m.text.strip()
    if u in ALLOWED_USERS or "*" in ALLOWED_USERS:
        await s.update_data(name=u, step=0, data=[], start=get_time())
        await m.answer("Введите название аптеки:")
        await s.set_state(Form.pharmacy)
    else:
        await m.answer("ФИО не распознано. Обратитесь в ИТ-отдел.")

@dp.message(Form.pharmacy)
async def proc_pharmacy(m: types.Message, s: FSMContext):
    await s.update_data(pharmacy=m.text.strip())
    await m.answer("Начинаем проверку…")
    await s.set_state(Form.rating)
    await send_quest(m.chat.id, s)

# --- CallbackQuery без фильтров ---
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, s: FSMContext):
    data = await s.get_data()
    step = data.get("step", 0)

    # обработка «Назад»
    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await s.set_data(data)
        await cb.answer("↩️ Назад")
        return await send_quest(cb.from_user.id, s)

    # обработка оценки вида score_X
    if cb.data and cb.data.startswith("score_"):
        sc = int(cb.data.split("_")[1])
        # сохраняем
        data.setdefault("data", []).append({"crit":criteria[step], "score":sc})
        data["step"] += 1
        await s.set_data(data)
        await cb.answer(f"✅ Вы выбрали {sc}")
        # правим пред. сообщение
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"Оценка: {sc} {'⭐'*sc}"
        )
        return await send_quest(cb.from_user.id, s)

    # всё остальное — игнор
    await cb.answer()

# --- Отправка вопроса ---
async def send_quest(chat_id: int, s: FSMContext):
    data = await s.get_data()
    step = data["step"]
    total = len(criteria)

    if step >= total:
        await bot.send_message(chat_id, "Проверка завершена. Формируем отчёт…")
        return await make_report(chat_id, data)

    c = criteria[step]
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

    await bot.send_message(chat_id, text, parse_mode=ParseMode.HTML, reply_markup=kb.as_markup())

# --- Генерация отчёта ---
async def make_report(chat_id: int, data):
    name  = data["name"]
    ts    = data["start"]
    pharm = data.get("pharmacy", "Без названия")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    title = (
        f"Отчёт по проверке аптеки\nИсполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells("A1:G2"); ws["A1"] = title; ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm
    headers = ["Блок","Критерий","Требование","Оценка участника","Макс. оценка","Примечание","Дата проверки"]
    for idx,h in enumerate(headers,1):
        cell = ws.cell(row=5, column=idx, value=h); cell.font = Font(bold=True)
    row = 6; total_sc=0; total_max=0
    for it in data["data"]:
        c = it["crit"]; sc = it["score"]
        ws.cell(row,1,c["block"]); ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"]); ws.cell(row,4,sc)
        ws.cell(row,5,c["max"]); ws.cell(row,7,ts)
        total_sc += sc; total_max += c["max"]; row += 1
    ws.cell(row+1,3,"ИТОГО:"); ws.cell(row+1,4,total_sc)
    ws.cell(row+2,3,"Максимум:"); ws.cell(row+2,4,total_max)

    fn = f"{pharm}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx".replace(" ","_")
    wb.save(fn)
    with open(fn,"rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f))
    os.remove(fn)
    log_csv(pharm, name, ts, total_sc, total_max)
    await bot.send_message(chat_id, "Готово! /start — чтобы начать заново.")

# --- Webhook & healthcheck ---
async def handle_webhook(request: web.Request):
    data = await request.json()
    upd = Update(**data)
    await dp.feed_update(bot=bot, update=upd)
    return web.Response(text="OK")

async def health(request: web.Request):
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True, allowed_updates=[])

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
