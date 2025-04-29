import logging
import os
import csv
import pytz
from datetime import datetime
from dotenv import load_dotenv

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiogram import Bot, Dispatcher, types, F
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile

# === Настройки ===
load_dotenv()
API_TOKEN = os.getenv("API_TOKEN")
if not API_TOKEN:
    raise RuntimeError("❌ Переменная API_TOKEN не задана!")

CHAT_ID = int(os.getenv("CHAT_ID", "0"))
if CHAT_ID == 0:
    raise RuntimeError("❌ Переменная CHAT_ID не задана!")

TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")

# === FSM-состояния ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Читаем чек-лист ===
df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение",
              "Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])

criteria = []
last_block = None
for _, r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({
        "block": blk,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# === Утилиты ===
def get_time():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(ph, nm, ts, sc, mx):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","ФИО проверяющего","Баллы","Макс"])
        w.writerow([ts, ph, nm, sc, mx])

# === Инициализация ===
logging.basicConfig(level=logging.INFO)
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp = Dispatcher(storage=MemoryStorage())

# --- /start ---
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n"
        "Интеллектуальная собственность ИТ Садыхан.\n\n"
        "📝 Заполняйте вдумчиво и внимательно, отчёт уйдёт департаменту качества.\n\n"
        "🏁 Чтобы начать, введите ваше ФИО:",
    )
    await state.set_state(Form.name)

# --- /id, /лог, /сброс ---
@dp.message(F.text == "/id")
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id: <code>{msg.chat.id}</code>")

@dp.message(F.text == "/лог")
async def cmd_log(msg: types.Message):
    if os.path.exists(LOG_PATH):
        await msg.answer_document(FSInputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не сформирован.")

@dp.message(F.text == "/сброс")
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Сброшено. /start — чтобы начать заново.")

# --- Получаем ФИО ---
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    user = msg.text.strip()
    await state.update_data(name=user, step=0, data=[], start=get_time())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# --- Название аптеки и старт ---
@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_quest(msg.chat.id, state)

# --- Кнопки оценки ---
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()  # мгновенный ACK
    data = await state.get_data()
    step = data.get("step", 0)

    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.set_data(data)
        return await send_quest(cb.from_user.id, state)

    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        data.setdefault("data", []).append({
            "crit": criteria[step],
            "score": score
        })
        data["step"] += 1
        await state.set_data(data)

        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_quest(cb.from_user.id, state)

# --- Шлём вопрос ---
async def send_quest(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total = len(criteria)

    if step >= total:
        await bot.send_message(chat_id, "✅ Проверка завершена. Формируем отчёт…")
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
    if step>0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# --- Формируем отчёт ---
async def make_report(chat_id: int, data):
    name     = data["name"]
    ts       = data["start"]
    pharm    = data.get("pharmacy","—")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    title = f"Отчет по проверке аптеки\nИсполнитель: {name}\nДата: {ts[:10]}"
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    headers = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата"]
    for i,h in enumerate(headers,1):
        ws.cell(row=5, column=i, value=h).font = Font(bold=True)

    row = 6; total_sc=0; total_mx=0
    for item in data["data"]:
        c = item["crit"]; sc=item["score"]
        ws.cell(row,1,c["block"])
        ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"])
        ws.cell(row,4,sc)
        ws.cell(row,5,c["max"])
        ws.cell(row,7,ts)
        total_sc+=sc; total_mx+=c["max"]; row+=1

    ws.cell(row+1,3,"ИТОГО:"); ws.cell(row+1,4,total_sc)
    ws.cell(row+2,3,"Максимум:"); ws.cell(row+2,4,total_mx)

    fname = f"{pharm}_{name}_{ts[:10]}.xlsx".replace(" ","_")
    wb.save(fname)

    # 1) в QA-чат
    with open(fname,"rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f))
    # 2) дубликат пользователю
    with open(fname,"rb") as f:
        await bot.send_document(chat_id, FSInputFile(f))
    os.remove(fname)

    log_csv(pharm,name,ts,total_sc,total_mx)
    await bot.send_message(chat_id, "📌 Отчёт отправлен в QA-чат. Чтобы ещё раз пройти — /start")

# === Запуск Polling ===
if __name__ == "__main__":
    logging.info("🚀 Запускаем polling…")
    dp.run_polling(bot)
