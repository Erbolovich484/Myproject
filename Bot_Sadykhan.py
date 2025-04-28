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
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")

ALLOWED_USERS = [
    "Николай Крылов",
    "Таждин Усейн",
    "Жанар Бөлтірік",
    "Шара Абдиева",
    "Тохтар Чарабасов",
    "*"
]

# === ЧТЕНИЕ КРИТЕРИЕВ ИЗ EXCEL ===
criteria_df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_index = criteria_df[criteria_df.iloc[:, 0] == "Блок"].index[0] + 1
criteria_df = criteria_df.iloc[start_index:, :8].reset_index(drop=True)
criteria_df.columns = [
    "Блок", "Критерий", "Требование", "Оценка",
    "Макс. значение", "Примечание", "Дата проверки", "Дата исправления"
]
criteria_df = criteria_df.dropna(subset=["Критерий", "Требование"])

criteria_list = []
last_block = None
for _, row in criteria_df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else last_block
    last_block = block
    max_val = (
        int(row["Макс. значение"])
        if pd.notna(row["Макс. значение"]) and str(row["Макс. значение"]).isdigit()
        else 10
    )
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
        writer = csv.writer(f)
        if not exists:
            writer.writerow(["Дата", "Аптека", "ФИО проверяющего", "Факт", "Макс. балл"])
        writer.writerow([timestamp, pharmacy, name, score, max_score])

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
    intro = (
        "📋 <b>Чек-лист посещения аптек</b>\n\n"
        "Заполняйте внимательно, отчет придет автоматически.\n\n"
        "🏁 Начнем!"
    )
    await message.answer(intro, parse_mode=ParseMode.HTML)
    await message.answer("Введите ваше ФИО:")
    await state.set_state(Form.name)

@router.message(F.text == "/лог")
async def send_log_file(message: types.Message):
    if os.path.exists(LOG_PATH):
        await message.answer_document(FSInputFile(LOG_PATH))
    else:
        await message.answer("Лог пока пуст.")

@router.message(Form.name)
async def process_name(message: types.Message, state: FSMContext):
    user_name = message.text.strip()
    if user_name in ALLOWED_USERS or "*" in ALLOWED_USERS:
        await state.update_data(name=user_name, step=0, data=[], start_time=get_astana_time())
        await message.answer("Введите название аптеки:")
        await state.set_state(Form.pharmacy)
    else:
        await message.answer("ФИО не распознано. Обратитесь в ИТ-отдел.")

@router.message(Form.pharmacy)
async def process_pharmacy(message: types.Message, state: FSMContext):
    await state.update_data(pharmacy=message.text.strip())
    await message.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_criterion(message.chat.id, state)

@router.callback_query(F.data.startswith("score_") | F.data == "prev")
async def process_score(callback: CallbackQuery, state: FSMContext):
    await callback.answer("✔️ Принято")
    data = await state.get_data()
    step = data.get('step', 0)

    # Обработка «Назад»
    if callback.data == "prev":
        if step > 0:
            data['step'] -= 1
            data['data'].pop()
            await state.set_data(data)
        return await send_criterion(callback.from_user.id, state)

    # Сохранение оценки
    score = int(callback.data.split("_")[1])
    if step < len(criteria_list):
        data.setdefault('data', []).append({"criterion": criteria_list[step], "score": score})
        data['step'] += 1
        await state.set_data(data)

    # Обновляем сообщение
    await bot.edit_message_text(
        chat_id=callback.message.chat.id,
        message_id=callback.message.message_id,
        text=f"✅ Оценка: {score} {'⭐'*score}"
    )
    await send_criterion(callback.from_user.id, state)

# === Функция отправки вопроса ===
async def send_criterion(chat_id, state: FSMContext):
    data = await state.get_data()
    step = data['step']
    total = len(criteria_list)

    # Если все вопросы пройдены
    if step >= total:
        await bot.send_message(chat_id, "Проверка завершена. Формируем отчёт…")
        await generate_and_send_excel(chat_id, data)
        await bot.send_message(chat_id, "Готово! Чтобы заново — /start")
        return await state.clear()

    c = criteria_list[step]

    # Формируем текст сообщения
    msg = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n\n"
        f"<b>Критерий:</b> {c['criterion']}\n\n"
        f"<b>Требование:</b> {c['requirement']}\n\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )

    # Строим клавиатуру: если max=1 — показываем 0 и 1, иначе от 1 до max
    kb = InlineKeyboardBuilder()
    start_score = 0 if c['max'] == 1 else 1
    for i in range(start_score, c['max'] + 1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(
        chat_id,
        msg,
        parse_mode=ParseMode.HTML,
        reply_markup=kb.as_markup()
    )

# === Генерация и отправка Excel-отчета ===
async def generate_and_send_excel(chat_id, session):
    name     = session['name']
    timestamp= session['start_time']
    pharmacy = session.get('pharmacy', 'Без названия')

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Заголовок
    title = (
        f"Отчёт по проверке аптеки\n"
        f"Исполнитель: {name}\n"
        f"Дата: {datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}\n"
        "Оформлен через бот"
    )
    ws.merge_cells('A1:G2')
    ws['A1'] = title
    ws['A1'].font = Font(size=14, bold=True)

    ws['B3'] = pharmacy

    # Шапка таблицы
    headers = ["Блок", "Критерий", "Требование", "Оценка участника", "Макс. оценка", "Примечание", "Дата проверки"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    # Заполняем данные
    row = 6
    total_score = 0
    total_max   = 0
    for item in session['data']:
        c = item['criterion']
        sc= item['score']
        ws.cell(row=row, column=1, value=c['block'])
        ws.cell(row=row, column=2, value=c['criterion'])
        ws.cell(row=row, column=3, value=c['requirement'])
        ws.cell(row=row, column=4, value=sc)
        ws.cell(row=row, column=5, value=c['max'])
        ws.cell(row=row, column=7, value=timestamp)
        total_score += sc
        total_max   += c['max']
        row += 1

    # Итого
    ws.cell(row=row+1, column=3, value="ИТОГО:")
    ws.cell(row=row+1, column=4, value=total_score)
    ws.cell(row=row+2, column=3, value="Максимум:")
    ws.cell(row=row+2, column=4, value=total_max)

    # Сохраняем файл
    date_str = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y")
    filename = f"{pharmacy}_{name}_{date_str}.xlsx".replace(" ", "_")
    wb.save(filename)

    # Отправляем в общий чат
    with open(filename, "rb") as f:
        await bot.send_document(CHAT_ID, types.BufferedInputFile(f.read(), filename=filename))

    # Локальный cleanup и лог
    os.remove(filename)
    log_checklist_submission(pharmacy, name, timestamp, total_score, total_max)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(dp.start_polling(bot))
