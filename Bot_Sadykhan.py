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

# === Настройки ===
load_dotenv()
API_TOKEN     = os.getenv("API_TOKEN")                      # Ваш токен
QA_CHAT_ID    = int(os.getenv("QA_CHAT_ID", "0"))           # чат для QA
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH= os.getenv("CHECKLIST_PATH", "checklist.xlsx")
LOG_PATH      = os.getenv("LOG_PATH", "checklist_log.csv")

# === Логирование ===
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === FSM-состояния ===
class Form(StatesGroup):
    name       = State()
    pharmacy   = State()
    rating     = State()
    comment    = State()

# === Чтение чек-листа ===
_df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = _df[_df.iloc[:,0]=="Блок"].index[0] + 1
_df = _df.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])
criteria = []
last_block = None
for _, r in _df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс"]) if pd.notna(r["Макс"]) and str(r["Макс"]).isdigit() else 10
    criteria.append({"block":blk,"criterion":r["Критерий"],"requirement":r["Требование"],"max":maxv})

TOTAL = len(criteria)

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

# === Инициализация ===
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=MemoryStorage())

# === /start ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Это интеллектуальная собственность ИТ-Садыхан.\n"
        "Заполняйте внимательно и вдумчиво — отчёт будет отправлен вам и в QA-чат.\n\n"
        "🏁 Введите ваше ФИО для авторизации:",
    )
    await state.set_state(Form.name)

# === ФИО ===
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    name = msg.text.strip()
    await state.update_data(name=name, step=0, answers=[], start=now_str("%Y-%m-%d %H:%M:%S"))
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# === Название аптеки ===
@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# === Кнопки оценок ===
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()  # всегда отвечаем, чтобы не "timeout"
    data = await state.get_data()
    step = data["step"]

    if cb.data == "prev" and step > 0:
        # назад
        data["step"] -= 1
        data["answers"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        crit  = criteria[step]
        data["answers"].append({"crit":crit,"score":score})
        data["step"] += 1
        await state.set_data(data)

        # редактируем предыдущее сообщение
        await bot.edit_message_text(
            f"✅ Оценка: {score} {'⭐'*score}",
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id
        )

        # если это последний — переходим к комменту
        if data["step"] >= TOTAL:
            await bot.send_message(cb.from_user.id, "✅ Проверка завершена. Добавьте, пожалуйста, вывод по аптеке (или «—», если нет):")
            return await state.set_state(Form.comment)

        return await send_question(cb.from_user.id, state)

# === Отправка вопроса ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    crit = criteria[step]
    kb = InlineKeyboardBuilder()
    start = 0 if crit["max"]==1 else 1
    for i in range(start, crit["max"]+1):
        kb.button(str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button("◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(
        chat_id,
        f"<b>Вопрос {step+1} из {TOTAL}</b>\n\n"
        f"<b>Блок:</b> {crit['block']}\n"
        f"<b>Критерий:</b> {crit['criterion']}\n"
        f"<b>Требование:</b> {crit['requirement']}\n"
        f"<b>Макс. балл:</b> {crit['max']}",
        reply_markup=kb.as_markup()
    )

# === Свободный комментарий ===
@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    await state.update_data(comment=msg.text.strip())
    await msg.answer("Формирую отчёт…")
    data = await state.get_data()
    await make_report(msg.chat.id, data)
    await state.clear()

# === Генерация и отправка отчёта ===
async def make_report(user_chat: int, data: dict):
    name     = data["name"]
    pharm    = data["pharmacy"]
    ts       = data["start"]
    comment  = data.get("comment", "")
    answers  = data["answers"]

    # Подготовка Excel
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    title = f"Отчёт: {pharm} — {name} ({ts.split()[0]})"
    ws.merge_cells("A1:G2")
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    ws["B3"] = pharm

    # Заголовки
    hdr = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата проверки"]
    for i,h in enumerate(hdr,1):
        ws.cell(row=5, column=i, value=h).font = Font(bold=True)

    # Данные
    row = 6; total=0; max_total=0
    for it in answers:
        c = it["crit"]; sc=it["score"]
        ws.cell(row,1,c["block"])
        ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"])
        ws.cell(row,4,sc)
        ws.cell(row,5,c["max"])
        ws.cell(row,7,ts)
        total += sc; max_total += c["max"]
        row += 1

    # ИТОГО и комментарий
    ws.cell(row+1,3,"ИТОГО:");   ws.cell(row+1,4,total)
    ws.cell(row+2,3,"Максимум:");ws.cell(row+2,4,max_total)
    ws.cell(row+4,1,"Вывод проверяющего:"); ws.cell(row+4,2, comment)

    # Сохранить файл
    fn = f"{pharm}_{name}_{now_str()}.xlsx".replace(" ", "_")
    wb.save(fn)

    # Отправка: пользователю и в QA-чат
    for chat in (user_chat, QA_CHAT_ID):
        with open(fn,"rb") as f:
            await bot.send_document(chat, types.InputFile(f, filename=fn))
    os.remove(fn)

    # Лог
    log_csv(pharm, name, ts, total, max_total)

# === Полезные команды ===
@dp.message(F.text == "/id")
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id = <code>{msg.chat.id}</code>")

@dp.message(F.text == "/лог")
async def cmd_log(msg: types.Message):
    if os.path.exists(LOG_PATH):
        await msg.answer_document(types.InputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не создан.")

@dp.message(F.text == "/сброс")
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново.")

# === Запуск ===
if __name__ == "__main__":
    logger.info("Старт polling…")
    dp.run_polling(bot)
