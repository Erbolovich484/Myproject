```python
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
from aiohttp import web

# === Загрузка конфигов ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")  # Telegram Bot API token
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))  # QA-чат ID
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
PORT           = int(os.getenv("PORT", 8000))

# === FSM-состояния ===
class Form(StatesGroup):
    name    = State()
    pharmacy= State()
    rating  = State()
    comment = State()

# === Чтение критериев из Excel ===
criteria_df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_idx = criteria_df[criteria_df.iloc[:,0] == "Блок"].index[0] + 1
criteria_df = criteria_df.iloc[start_idx:, :8].reset_index(drop=True)
criteria_df.columns = [
    "Блок", "Критерий", "Требование", "Оценка",
    "Макс. значение", "Примечание", "Дата проверки", "Дата исправления"
]
criteria_df = criteria_df.dropna(subset=["Критерий", "Требование"]).
                reset_index(drop=True)

criteria = []
last_block = None
for _, row in criteria_df.iterrows():
    block = row["Блок"] if pd.notna(row["Блок"]) else last_block
    last_block = block
    maxv = int(row["Макс. значение"]) if pd.notna(row["Макс. значение"]) and str(row["Макс. значение"]).isdigit() else 10
    criteria.append({
        "block": block,
        "criterion": row["Критерий"],
        "requirement": row["Требование"],
        "max": maxv
    })

def current_time():
    tz = pytz.timezone("Asia/Almaty")
    return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")

# === Logging setup ===
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

# === Bot & Dispatcher ===
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp = Dispatcher(bot=bot, storage=MemoryStorage())

# === Handlers ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    logger.info("CMD /start from %s", msg.from_user.id)
    await state.clear()
    await msg.answer(
        "<b>📋 Чек‑лист посещения аптек</b>\n"
        "<i>Интеллектуальная собственность ИТ «Садыхан»</i>\n\n"
        "Заполняйте чек‑лист вдумчиво:\n"
        "— Inline‑кнопки для быстрой оценки;\n"
        "— После всех оценок введите итоговый вывод.\n\n"
        "Введите ваше ФИО:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)

@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    await state.update_data(
        name=msg.text.strip(),
        step=0,
        data=[],
        start=current_time()
    )
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data.get("step", 0)
    total = len(criteria)
    logger.info("send_question: step=%d/%d to %s", step, total, chat_id)
    if step >= total:
        await bot.send_message(
            chat_id,
            "✅ Оценки собраны! Теперь введите ваши выводы по аптеке:",
        )
        await state.set_state(Form.comment)
        return

    c = criteria[step]
    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )
    kb = InlineKeyboardBuilder()
    for i in range(1, c['max'] + 1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(
        chat_id,
        text,
        reply_markup=kb.as_markup()
    )

@dp.callback_query()
async def cb_handler(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()
    data = await state.get_data()
    step = data.get("step", 0)
    if cb.data == "prev" and step > 0:
        data['step'] -= 1
        data['data'].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    if cb.data and cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        if step < len(criteria):
            data.setdefault('data', []).append({
                'crit': criteria[step],
                'score': score
            })
            data['step'] += 1
            await state.set_data(data)
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    data = await state.get_data()
    data['comment'] = msg.text.strip()
    await state.set_data(data)
    await msg.answer("Формируем отчёт…")
    await generate_and_send(msg.chat.id, data)

async def generate_and_send(chat_id: int, data: dict):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Шапка отчёта
    title = (
        f"Отчёт по проверке: {data['pharmacy']}\n"
        f"Исполнитель: {data['name']}\n"
        f"Дата: {datetime.strptime(data['start'], '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells('A1:G2')
    ws['A1'] = title
    ws['A1'].font = Font(size=14, bold=True)
    ws['B3'] = data['pharmacy']

    headers = ["Блок","Критерий","Требование","Баллы","Макс","Дата" ]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=idx, value=h)
        cell.font = Font(bold=True)

    row = 6
    total_score = 0
    total_max = 0
    for item in data['data']:
        c = item['crit']
        sc = item['score']
        ws.cell(row,1,c['block'])
        ws.cell(row,2,c['criterion'])
        ws.cell(row,3,c['requirement'])
        ws.cell(row,4,sc)
        ws.cell(row,5,c['max'])
        ws.cell(row,6,data['start'])
        total_score += sc
        total_max += c['max']
        row += 1

    ws.cell(row+1,3,"ИТОГО:")
    ws.cell(row+1,4,total_score)
    ws.cell(row+2,3,"Максимум:")
    ws.cell(row+2,4,total_max)
    # Вывод проверяющего
    ws.merge_cells(start_row=row+4, start_column=1, end_row=row+4, end_column=7)
    ws.cell(row+4,1, f"Вывод проверяющего: {data['comment']}")
    ws.cell(row+4,1).font = Font(bold=True)

    fname = f"{data['pharmacy']}_{data['name']}_{datetime.strptime(data['start'], '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx".replace(' ','_')
    wb.save(fname)

    # Отправка
    for target in (chat_id, CHAT_ID):
        with open(fname, 'rb') as f:
            await bot.send_document(target, FSInputFile(f, filename=fname))
    os.remove(fname)

# === Webhook & server ===
async def handle_update(request: web.Request):
    data = await request.json()
    upd = Update(**data)
    logger.info("Incoming raw update: %s", data)
    await dp.feed_update(bot, upd)
    return web.Response(text='OK')

app = web.Application()
app.router.add_post(f"/webhook/{API_TOKEN}", handle_update)

if __name__ == '__main__':
    logger.info("Starting server on port %d", PORT)
    web.run_app(app, host='0.0.0.0', port=PORT)
```
