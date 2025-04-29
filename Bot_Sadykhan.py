import logging
import os
import csv
import pytz
from datetime import datetime
from dotenv import load_dotenv

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiogram import Bot, Dispatcher, types
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update

from aiohttp import web

# ========== Настройки ==========
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")      # ваш токен
QA_CHAT_ID     = int(os.getenv("CHAT_ID"))   # chat_id QA-чата
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
PORT           = int(os.getenv("PORT", "8000"))

logging.basicConfig(level=logging.INFO)

# ========== FSM ==========
class Form(StatesGroup):
    name       = State()
    pharmacy   = State()
    rating     = State()
    conclusion = State()   # ← новый шаг

# ========== Чтение чек-листа ==========
df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])

criteria = []
last_block = None
for _, r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({"block": blk, "criterion": r["Критерий"], "requirement": r["Требование"], "max": maxv})

def get_time():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy, name, ts, score, max_score):
    path = "checklist_log.csv"
    exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# ========== Инициализация бота ==========
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=MemoryStorage())

# ========== Хэндлеры команд ==========
@dp.message(commands=["start"])
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n"
        "💡 Интеллектуальная собственность ИТ-отдела «Садыхан».\n\n"
        "✍️ Заполняйте вдумчиво и неспеша — кнопки активны всегда.\n"
        "✅ По завершении отчёт придёт в Excel в QA-чат и вам.\n\n"
        "🏁 Введите ваше ФИО:",
    )
    await state.set_state(Form.name)

@dp.message(commands=["id"])
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id: <code>{msg.chat.id}</code>")

# ========== Авторизация и начало ==========
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    name = msg.text.strip()
    await state.update_data(name=name, step=0, data=[], start=get_time())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_quest(msg.chat.id, state)

# ========== Обработка inline-кнопок ==========
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()
    data = await state.get_data()
    step = data.get("step", 0)

    # Назад
    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.set_data(data)
        return await send_quest(cb.from_user.id, state)

    # Оценка
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        data.setdefault("data", []).append({"crit": criteria[step], "score": score})
        data["step"] += 1
        await state.set_data(data)

        # Редактируем сообщение с оценкой
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )

        # Если вопросы закончились — идём в conclusion
        if data["step"] >= len(criteria):
            await bot.send_message(cb.from_user.id,
                "📝 Проверка завершена. Пожалуйста, напишите свои выводы по аптеке:")
            return await state.set_state(Form.conclusion)

        return await send_quest(cb.from_user.id, state)

# ========== Отправка следующего вопроса ==========
async def send_quest(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)

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
        kb.button(str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button("◀️ Назад", callback_data="prev")
    kb.adjust(5)
    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# ========== Обработка вывода ==========
@dp.message(Form.conclusion)
async def proc_conclusion(msg: types.Message, state: FSMContext):
    await state.update_data(conclusion=msg.text.strip())
    data = await state.get_data()
    await msg.answer("✅ Спасибо! Формируем отчёт…")
    await make_report(msg.chat.id, data)
    await state.clear()

# ========== Генерация и отправка отчёта ==========
async def make_report(chat_id: int, data):
    name       = data["name"]
    ts         = data["start"]
    pharmacy   = data.get("pharmacy","—")
    conclusion = data.get("conclusion","")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Шапка
    title = (f"Отчёт по проверке аптеки\n"
             f"Исполнитель: {name}\n"
             f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}")
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharmacy

    # Заголовки колонок
    headers = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки"]
    for idx,h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    # Заполняем строки
    row = 6
    total_sc = total_mx = 0
    for item in data["data"]:
        c  = item["crit"]
        sc = item["score"]
        ws.cell(row,1,c["block"])
        ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"])
        ws.cell(row,4,sc)
        ws.cell(row,5,c["max"])
        ws.cell(row,7,ts)
        total_sc += sc
        total_mx += c["max"]
        row += 1

    # Итого
    ws.cell(row+1,3,"ИТОГО:")
    ws.cell(row+1,4,total_sc)
    ws.cell(row+2,3,"Максимум:")
    ws.cell(row+2,4,total_mx)

    # Выводы проверяющего
    ws.cell(row+4,1,"Выводы проверяющего:")
    ws.merge_cells(start_row=row+4, start_column=2, end_row=row+4, end_column=6)
    ws.cell(row+4,2,conclusion)

    # Сохраняем и отправляем
    fname = f"{pharmacy}_{name}_{ts[:10]}.xlsx".replace(" ","_")
    wb.save(fname)
    with open(fname,"rb") as f:
        await bot.send_document(QA_CHAT_ID, FSInputFile(f))
    with open(fname,"rb") as f:
        await bot.send_document(chat_id, FSInputFile(f))
    os.remove(fname)

    log_csv(pharmacy, name, ts, total_sc, total_mx)

    # Финальное сообщение
    await bot.send_message(chat_id,
        "📌 Отчёт отправлен в QA-чат и вам.\n"
        "Чтобы пройти ещё раз — /start")

# ========== Webhook & запуск сервера ==========
async def handle_webhook(request: web.Request):
    data = await request.json()
    upd  = Update(**data)
    await dp.feed_update(bot, upd)
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    # если нужен вебхук, укажите свой URL:
    # await bot.set_webhook(os.getenv("WEBHOOK_URL"), drop_pending_updates=True)
    pass

app = web.Application()
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(on_startup)

if __name__ == "__main__":
    web.run_app(app, host="0.0.0.0", port=PORT)
