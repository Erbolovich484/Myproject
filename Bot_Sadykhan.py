import logging
import os
import csv
import pytz
from datetime import datetime
from dotenv import load_dotenv

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiohttp import web
from aiogram import Bot, Dispatcher, types
from aiogram.enums import ParseMode
from aiogram.filters import Command
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update

# === Загружаем .env ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")  # Например https://<your-domain>/webhook
PORT           = int(os.getenv("PORT", "8080"))

# === FSM ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()
    comment  = State()

# === Читаем критерии из Excel ===
_df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = _df[_df.iloc[:,0]=="Блок"].index[0] + 1
_df = _df.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])
criteria = []
_last_blk = None
for _, r in _df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else _last_blk
    _last_blk = blk
    maxv = int(r["Макс"]) if pd.notna(r["Макс"]) and str(r["Макс"]).isdigit() else 10
    criteria.append({
        "block": blk,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# === Вспомогательные функции ===
def now_str():
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy, name, ts, score, max_score):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","Проверяющий","Баллы","Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# === Инициализация бота и диспетчера ===
logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=MemoryStorage())

# === Хэндлеры команд ===
@dp.message(Command("start"))
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n"
        "🧠 Интеллектуальная собственность ИТ «Садыхан»\n\n"
        "Заполняйте вдумчиво и внимательно.\n"
        "По завершении отчёт придёт в QA-чат и вам в личку.\n\n"
        "🏁 Введите своё ФИО для авторизации:",
    )
    await state.set_state(Form.name)

@dp.message(Command("id"))
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id: <code>{msg.chat.id}</code>")

@dp.message(Command("лог"))
async def cmd_log(msg: types.Message):
    if os.path.exists(LOG_PATH):
        await msg.answer_document(FSInputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не сформирован.")

@dp.message(Command("сброс"))
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Состояние сброшено. /start — чтобы начать заново.")

# === Авторизация и ввод аптеки ===
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    user = msg.text.strip()
    await state.update_data(name=user, step=0, answers=[], start=now_str())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# === Обработка нажатий inline-кнопок ===
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()  # устраняем «query is too old»
    data = await state.get_data()
    step = data.get("step", 0)

    # если оценка
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        data["answers"].append({"crit": criteria[step], "score": score})
        data["step"] += 1
        await state.set_data(data)
        # редактируем сообщение
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

    # «назад»
    if cb.data == "prev" and step>0:
        data["step"] -= 1
        data["answers"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

# === Отправка вопроса ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)

    if step >= total:
        await bot.send_message(chat_id, "✅ Проверка завершена. Формируем отчёт…")
        await state.set_state(Form.comment)
        return await bot.send_message(chat_id, "✍️ Напишите, пожалуйста, свои выводы по аптеке:")

    c = criteria[step]
    kb = InlineKeyboardBuilder()
    start = 0 if c["max"]==1 else 1
    for i in range(start, c["max"]+1):
        kb.button(str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button("◀️ Назад", callback_data="prev")
    kb.adjust(5)

    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )
    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# === Сбор дополнительного комментария + отчёт ===
@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    data       = await state.get_data()
    comment    = msg.text.strip()
    data["comment"] = comment

    # строим Excel
    wb   = load_workbook(TEMPLATE_PATH)
    ws   = wb.active
    ts   = data["start"]
    name = data["name"]
    pharm = data["pharmacy"]

    # шапка
    ws.merge_cells("A1:G2")
    ws["A1"] = f"Отчёт по аптеке: {pharm}\nПроверил: {name}\nДата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"]    = pharm

    # заголовки таблицы
    headers = ["Блок","Критерий","Требование","Оценка","Макс","Комментарий","Дата"]
    for i,h in enumerate(headers,1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.font = Font(bold=True)

    row = 6
    total_sc = total_max = 0
    for item in data["answers"]:
        c = item["crit"]
        sc = item["score"]
        ws.cell(row,1, c["block"])
        ws.cell(row,2, c["criterion"])
        ws.cell(row,3, c["requirement"])
        ws.cell(row,4, sc)
        ws.cell(row,5, c["max"])
        ws.cell(row,6, "")        # пусто под построчный комментарий
        ws.cell(row,7, ts)
        total_sc += sc
        total_max += c["max"]
        row += 1

    # итоги и общий комментарий
    ws.cell(row+1,3,"ИТОГО:")
    ws.cell(row+1,4,total_sc)
    ws.cell(row+2,3,"Максимум:")
    ws.cell(row+2,4,total_max)
    ws.cell(row+4,1,"Выводы проверяющего:")
    ws.merge_cells(start_row=row+4, start_column=1, end_row=row+8, end_column=7)
    ws.cell(row+4,1, data["comment"])

    # сохраняем и отправляем
    fn = f"{pharm}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx"
    wb.save(fn)
    with open(fn, "rb") as f:
        # в QA-чат
        await bot.send_document(QA_CHAT_ID, FSInputFile(f, fn))
    with open(fn, "rb") as f:
        # дублируем пользователю
        await bot.send_document(msg.chat.id, FSInputFile(f, fn))
    os.remove(fn)

    # логируем
    log_csv(pharm, name, ts, total_sc, total_max)
    await msg.answer("✅ Отчёт готов и отправлен.\nЧтобы начать заново — /start")
    await state.clear()

# === Webhook & запуск сервера ===
async def handle_webhook(req: web.Request):
    data = await req.json()
    upd  = Update(**data)
    await dp.feed_update(bot, upd)
    return web.Response(text="OK")

async def on_startup(app: web.Application):
    logging.info(f"Устанавливаю Webhook: {WEBHOOK_URL}")
    await bot.set_webhook(WEBHOOK_URL)

async def on_cleanup(app: web.Application):
    await bot.delete_webhook()

app = web.Application()
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__ == "__main__":
    logging.info(f"Старт сервера на порту {PORT}")
    web.run_app(app, host="0.0.0.0", port=PORT)
