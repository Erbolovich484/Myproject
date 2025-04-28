import logging
import pytz
from datetime import datetime
import asyncio
import os
import csv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F, Router, types
from aiogram.enums import ParseMode
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.client.default import DefaultBotProperties

# === ЗАГРУЗКА ПЕРЕМЕННЫХ ОКРУЖЕНИЯ ===
load_dotenv()
API_TOKEN     = os.getenv("API_TOKEN")
CHAT_ID       = int(os.getenv("CHAT_ID"))
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH= os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH      = os.getenv("LOG_PATH", "checklist_log.csv")

ALLOWED_USERS = [
    "Николай Крылов", "Таждин Усейн",
    "Жанар Бөлтірік", "Шара Абдиева",
    "Тохтар Чарабасов", "*"
]

# === ЧТЕНИЕ КРИТЕРИЕВ ИЗ EXCEL ===
criteria_df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_index = criteria_df[criteria_df.iloc[:, 0] == "Блок"].index[0] + 1
criteria_df = criteria_df.iloc[start_index:, :8].reset_index(drop=True)
criteria_df.columns = [
    "Блок","Критерий","Требование","Оценка",
    "Макс. значение","Примечание","Дата проверки","Дата исправления"
]
criteria_df = criteria_df.dropna(subset=["Критерий","Требование"])

criteria_list = []
last_block = None
for _, row in criteria_df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else last_block
    last_block = block
    max_val = int(row["Макс. значение"]) if pd.notna(row["Макс. значение"]) and str(row["Макс. значение"]).isdigit() else 10
    criteria_list.append({
        "block": block,
        "criterion": row["Критерий"],
        "requirement": row["Требование"],
        "max": max_val
    })

# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===
def get_astana_time():
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")

def log_checklist_submission(pharmacy, name, timestamp, score, max_score):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","ФИО проверяющего","Факт","Макс. балл"])
        w.writerow([timestamp, pharmacy, name, score, max_score])

# === FSM СТЕЙТЫ ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === ИНИЦИАЛИЗАЦИЯ БОТА ===
bot = Bot(token=API_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
storage = MemoryStorage()
dp = Dispatcher(storage=storage)
router = Router()
dp.include_router(router)

# === ХАНДЛЕРЫ ===

@router.message(F.text == "/start")
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    text = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Заполняйте внимательно. Отчёт придёт автоматически.\n\n"
        "🏁 Начнём!"
    )
    await message.answer(text)
    await message.answer("Введите ваше ФИО:")
    await state.set_state(Form.name)

@router.message(F.text == "/лог")
async def send_log_file(message: types.Message):
    if os.path.exists(LOG_PATH):
        await message.answer_document(FSInputFile(LOG_PATH))
    else:
        await message.answer("Лог пока пустой.")

@router.message(Form.name)
async def process_name(message: types.Message, state: FSMContext):
    name = message.text.strip()
    if name in ALLOWED_USERS or "*" in ALLOWED_USERS:
        await state.update_data(name=name, step=0, data=[], start_time=get_astana_time())
        await message.answer("Введите название аптеки:")
        await state.set_state(Form.pharmacy)
    else:
        await message.answer("ФИО не распознано.")

@router.message(Form.pharmacy)
async def process_pharmacy(message: types.Message, state: FSMContext):
    await state.update_data(pharmacy=message.text.strip())
    await message.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_criterion(message.chat.id, state)

@router.callback_query(F.data.startswith("score_") | (F.data == "prev"))
async def process_score(cb: CallbackQuery, state: FSMContext):
    await cb.answer("✔️")
    data = await state.get_data()
    step = data['step']
    if cb.data == "prev" and step > 0:
        data['step'] -= 1
        data['data'].pop()
        await state.set_data(data)
        return await send_criterion(cb.from_user.id, state)

    score = int(cb.data.split("_")[1])
    data.setdefault('data',[]).append({"criterion": criteria_list[step], "score": score})
    data['step'] += 1
    await state.set_data(data)

    # Обновляем текст кнопки
    await bot.edit_message_text(
        chat_id=cb.message.chat.id,
        message_id=cb.message.message_id,
        text=f"✅ Оценка: {score} {'⭐'*score}"
    )
    await send_criterion(cb.from_user.id, state)

async def send_criterion(chat_id, state: FSMContext):
    data = await state.get_data()
    step = data['step']
    if step >= len(criteria_list):
        await bot.send_message(chat_id, "Готовим отчёт…")
        await generate_and_send_excel(chat_id, data)
        await bot.send_message(chat_id, "Готово! /start — чтобы заново.")
        return await state.clear()

    c = criteria_list[step]
    kb = InlineKeyboardBuilder()
    for i in range(1, c['max']+1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    txt = (
        f"<b>Вопрос {step+1} из {len(criteria_list)}</b>\n"
        f"<b>{c['block']}</b>\n"
        f"{c['criterion']}\n"
        f"Макс: {c['max']}"
    )
    await bot.send_message(chat_id, txt, reply_markup=kb.as_markup(), parse_mode=ParseMode.HTML)

async def generate_and_send_excel(chat_id, session):
    name     = session['name']
    ts       = session['start_time']
    pharmacy = session.get('pharmacy','Без названия')

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Заголовок
    title = (f"Отчёт по проверке аптеки\nИсполнитель: {name}\n"
             f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}\n"
             "Через бот")
    ws.merge_cells('A1:G2')
    ws['A1'] = title
    ws['A1'].font = Font(size=14, bold=True)

    ws['B3'] = pharmacy
    # Шапка таблицы
    headers = ["Блок","Критерий","Требование","Оценка участника","Макс оценка","Примечание","Дата"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    # Заполняем данные
    row = 6
    total, total_max = 0, 0
    for item in session['data']:
        c = item['criterion']
        sc= item['score']
        ws.cell(row=row, column=1, value=c['block'])
        ws.cell(row=row, column=2, value=c['criterion'])
        ws.cell(row=row, column=3, value=c['requirement'])
        ws.cell(row=row, column=4, value=sc)
        ws.cell(row=row, column=5, value=c['max'])
        ws.cell(row=row, column=7, value=ts)
        total += sc; total_max += c['max']
        row += 1

    # Итого
    ws.cell(row=row+1, column=3, value="ИТОГО:")
    ws.cell(row=row+1, column=4, value=total)
    ws.cell(row=row+2, column=3, value="Максимум:")
    ws.cell(row=row+2, column=4, value=total_max)

    # Сохранить и отправить
    date_str = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y")
    fname = f"{pharmacy}_{name}_{date_str}.xlsx".replace(" ", "_")
    wb.save(fname)

    with open(fname, "rb") as f:
        await bot.send_document(CHAT_ID, types.BufferedInputFile(f.read(), filename=fname))

    os.remove(fname)
    log_checklist_submission(pharmacy, name, ts, total, total_max)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(dp.start_polling(bot))
