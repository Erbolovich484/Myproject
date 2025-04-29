import os
import csv
import logging
import pytz

from datetime import datetime
from dotenv import load_dotenv

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiohttp import web
from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ParseMode
from aiogram.types import FSInputFile, Update
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.default import DefaultBotProperties

# === Загрузка переменных окружения (.env или Railway Variables) ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))            # ваш QA-чат
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")                   # например https://<ваш-сервис>.railway.app/webhook
PORT           = int(os.getenv("PORT", "8000"))

# === FSM-состояния ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Чтение критериев из Excel ===
_df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_idx = _df[_df.iloc[:,0] == "Блок"].index[0] + 1
_df = _df.iloc[start_idx:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])

criteria = []
_last_block = None
for _, row in _df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else _last_block
    _last_block = block
    maxv = int(row["Макс. значение"]) if pd.notna(row["Макс. значение"]) and str(row["Макс. значение"]).isdigit() else 10
    criteria.append({
        "block":       block,
        "criterion":   row["Критерий"],
        "requirement": row["Требование"],
        "max":         maxv
    })

# === Логирование ===
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def get_time():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy, name, ts, score, max_score):
    header = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if header:
            w.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# === Инициализация бота и диспетчера ===
session = AiohttpSession()
bot = Bot(
    token=API_TOKEN,
    session=session,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML)
)
dp = Dispatcher(bot=bot, storage=MemoryStorage())

# === Команды ===
@dp.message(F.text == "/start")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    intro = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Интеллектуальная собственность ИТ-отдела «Садыхан».\n"
        "Пожалуйста, заполняйте вдумчиво и внимательно.\n"
        "По завершении отчёт автоматически отправится в QA-чат.\n\n"
        "🏁 Введите ваше ФИО для авторизации:"
    )
    await message.answer(intro)
    await state.set_state(Form.name)

@dp.message(F.text == "/id")
async def cmd_id(message: types.Message):
    await message.answer(f"Ваш chat_id: <code>{message.chat.id}</code>")

@dp.message(F.text == "/лог")
async def cmd_log(message: types.Message):
    if os.path.exists(LOG_PATH):
        await message.answer_document(FSInputFile(LOG_PATH))
    else:
        await message.answer("Лог пока пуст.")

@dp.message(F.text == "/сброс")
async def cmd_reset(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer("Состояние сброшено. Чтобы начать заново — /start")

# === Сбор данных ФИО и аптеки ===
@dp.message(Form.name)
async def process_name(message: types.Message, state: FSMContext):
    name = message.text.strip()
    await state.update_data(name=name, step=0, data=[], start=get_time())
    await message.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def process_pharmacy(message: types.Message, state: FSMContext):
    await state.update_data(pharmacy=message.text.strip())
    await message.answer("✅ Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(message.chat.id, state)

# === Обработка inline-кнопок ===
@dp.callback_query()
async def callback_handler(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()  # мгновенный ACK
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
            data.setdefault("data", []).append({
                "crit":  criteria[step],
                "score": score
            })
            data["step"] += 1
            await state.set_data(data)

        # редактируем предыдущее сообщение
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

# === Функция отправки следующего вопроса ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)

    # Если вопросов больше нет — отчёт
    if step >= total:
        await bot.send_message(chat_id, "🎉 Проверка завершена. Формируем отчёт…")
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
    kb.adjust(min(c["max"]+1, 5))

    await bot.send_message(chat_id, text,
                           parse_mode=ParseMode.HTML,
                           reply_markup=kb.as_markup())

# === Формирование и отправка отчёта ===
async def make_report(chat_id: int, data: dict):
    name     = data["name"]
    ts       = data["start"]
    pharmacy = data.get("pharmacy", "Без названия")

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
    ws["B3"] = pharmacy

    headers = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    row = 6
    total_score = 0
    total_max   = 0
    for item in data["data"]:
        c  = item["crit"]
        sc = item["score"]
        ws.cell(row, 1, c["block"])
        ws.cell(row, 2, c["criterion"])
        ws.cell(row, 3, c["requirement"])
        ws.cell(row, 4, sc)
        ws.cell(row, 5, c["max"])
        ws.cell(row, 7, ts)
        total_score += sc
        total_max   += c["max"]
        row += 1

    ws.cell(row+1, 3, "ИТОГО:")
    ws.cell(row+1, 4, total_score)
    ws.cell(row+2, 3, "Максимум:")
    ws.cell(row+2, 4, total_max)

    filename = f"{pharmacy}_{name}_{ts[:10].replace('-','_')}.xlsx".replace(" ", "_")
    wb.save(filename)

    # 1) в QA-чат
    with open(filename, "rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f, filename))
    # 2) пользователю
    with open(filename, "rb") as f:
        await bot.send_document(chat_id, FSInputFile(f, filename))
    os.remove(filename)
    log_csv(pharmacy, name, ts, total_score, total_max)

    await bot.send_message(chat_id,
        "✅ Отчёт готов и отправлен в QA-чат.\n"
        "Чтобы пройти еще раз — /start"
    )

# === Webhook & healthcheck для Railway ===
async def handle_webhook(request: web.Request):
    payload = await request.json()
    upd     = Update(**payload)
    await dp.feed_update(update=upd, bot=bot)
    return web.Response(text="OK")

async def health(request: web.Request):
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    logger.info("Устанавливаю Webhook: %s", WEBHOOK_URL)
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True)

async def on_cleanup(app: web.Application):
    await bot.delete_webhook()
    await dp.storage.close()

app = web.Application()
app.router.add_get("/", health)
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__ == "__main__":
    logger.info("Старт сервера на порту %s…", PORT)
    web.run_app(app, host="0.0.0.0", port=PORT)
