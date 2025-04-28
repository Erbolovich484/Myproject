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
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.default import DefaultBotProperties

from aiohttp import web

# === ЗАГруЗКА ОКРУЖЕНИЯ ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")      # https://<app>.onrender.com/webhook
PORT           = int(os.getenv("PORT", "8000"))

# === FSM STATES ===
class Form(StatesGroup):
    name       = State()
    pharmacy   = State()
    rating     = State()
    conclusion = State()

# === ЧитАем критерии ===
_df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = _df[_df.iloc[:,0]=="Блок"].index[0] + 1
_df = _df.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])

criteria = []
_last_block = None
for _, r in _df.iterrows():
    block = r["Блок"] if pd.notna(r["Блок"]) else _last_block
    _last_block = block
    maxv = int(r["Макс"]) if pd.notna(r["Макс"]) and str(r["Макс"]).isdigit() else 10
    criteria.append({
        "block": block,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# === УтИЛИТЫ ===
def now_ts() -> str:
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy, name, ts, score, max_score):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
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

# === СТАРТ и Приветствие ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    welcome = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "© Интеллектуальная собственность ИТ-службы «Садыхан».\n"
        "Заполняйте вдумчиво — отчёт автоматически уйдёт в QA-чат.\n\n"
        "🏁 Для старта введите ваше ФИО:"
    )
    await msg.answer(welcome)
    await state.set_state(Form.name)

# === Название аптеки ===
@dp.message(Form.name)
async def name_handler(msg: types.Message, state: FSMContext):
    await state.update_data(name=msg.text.strip(), step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# === Начинаем проверку ===
@dp.message(Form.pharmacy)
async def pharm_handler(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# === Обработка inline-кнопок ===
@dp.callback_query()
async def cb_handler(cb: types.CallbackQuery, state: FSMContext):
    # сразу подтверждаем callback, чтобы кнопка осталась активной
    await cb.answer()

    data = await state.get_data()
    step = data.get("step", 0)

    # Назад
    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    # Оценка
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        if step < len(criteria):
            data.setdefault("data", []).append({"crit": criteria[step], "score": score})
            data["step"] += 1
            await state.set_data(data)

        # редактируем прошлое сообщение
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

# === Отправка следующего вопроса ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)

    # если всё пройдено
    if step >= total:
        await bot.send_message(chat_id, "✅ Проверка завершена. Пожалуйста, введите ваш вывод по аптеке:")
        return await state.set_state(Form.conclusion)

    c = criteria[step]
    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )

    # inline-кнопки
    kb = InlineKeyboardBuilder()
    start = 0 if c["max"] == 1 else 1
    for i in range(start, c["max"]+1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, parse_mode=ParseMode.HTML, reply_markup=kb.as_markup())

# === Сбор вывода и генерация отчёта ===
@dp.message(Form.conclusion)
async def conclusion_handler(msg: types.Message, state: FSMContext):
    await state.update_data(conclusion=msg.text.strip())
    data = await state.get_data()

    # готовим Excel
    ts    = data["start"]
    name  = data["name"]
    pharm = data["pharmacy"]
    wb    = load_workbook(TEMPLATE_PATH)
    ws    = wb.active

    # заголовок
    title = (
        f"Отчёт по проверке аптеки\n"
        f"Исполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    # таблица
    headers = ["Блок","Критерий","Требование","Баллы","Макс","Дата проверки"]
    for idx,h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    row = 6
    total_score = 0
    total_max   = 0
    for item in data["data"]:
        c = item["crit"]
        sc = item["score"]
        ws.cell(row,1,c["block"])
        ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"])
        ws.cell(row,4,sc)
        ws.cell(row,5,c["max"])
        ws.cell(row,6,ts)
        total_score += sc
        total_max   += c["max"]
        row += 1

    # итоги
    ws.cell(row+1,3,"ИТОГО:")
    ws.cell(row+1,4,total_score)
    ws.cell(row+2,3,"Максимум:")
    ws.cell(row+2,4,total_max)

    # вывод по аптеке
    ws.cell(row+4,1,"Вывод аудитора:")
    ws.cell(row+4,2,data["conclusion"])

    # сохраняем
    fn = f"{pharm}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d%m%Y')}.xlsx".replace(" ","_")
    wb.save(fn)

    # отправляем в QA-чат и пользователю
    with open(fn,"rb") as f:
        await bot.send_document(QA_CHAT_ID, FSInputFile(f, filename=fn))
    with open(fn,"rb") as f:
        await bot.send_document(msg.chat.id, FSInputFile(f, filename=fn))

    os.remove(fn)
    log_csv(pharm, name, ts, total_score, total_max)

    # финальное сообщение
    await msg.answer("🎉 Отчёт сформирован и отправлен в QA-чат.\nЧтобы пройти заново — /start")
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
