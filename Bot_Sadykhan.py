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
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.default import DefaultBotProperties

from aiohttp import web

# === Загрузка конфигурации ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
QA_CHAT_ID     = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")
PORT           = int(os.getenv("PORT", "8000"))

# === FSM-состояния ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Чтение чек-листа ===
df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0] == "Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"]);
criteria=[]; last_block=None
for _,r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({"block":blk,"criterion":r["Критерий"],"requirement":r["Требование"],"max":maxv})

# === Утилиты ===
def now_ts():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(ph, nm, ts, sc, mx):
    new = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if new: w.writerow(["Дата","Аптека","ФИО","Баллы","Макс"])
        w.writerow([ts, ph, nm, sc, mx])

# === Инициализация бота ===
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session,
          default=DefaultBotProperties(parse_mode=ParseMode.HTML))
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# === Хэндлеры ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n" \
        "Введите ваше ФИО для авторизации:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)

@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    await state.update_data(name=msg.text.strip(), step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

@dp.callback_query(F.data.startswith("score_") | F.data == "prev")
async def cb_handler(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()
    data = await state.get_data(); step = data.get("step",0)
    # Назад
    if cb.data == "prev" and step > 0:
        data["step"] -= 1; data["data"].pop(); await state.set_data(data)
        return await send_question(cb.from_user.id, state)
    # Оценка
    score = int(cb.data.split("_")[1])
    if step < len(criteria):
        data.setdefault("data",[]).append({"crit":criteria[step],"score":score})
        data["step"]+=1; await state.set_data(data)
    # Редактируем текст
    await bot.edit_message_text(
        chat_id=cb.message.chat.id,
        message_id=cb.message.message_id,
        text=f"✅ Оценка: {score} {'⭐'*score}"
    )
    return await send_question(cb.from_user.id, state)

async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data(); step = data["step"]; total = len(criteria)
    # Завершение
    if step >= total:
        await bot.send_message(chat_id, "Проверка завершена. Формируем отчёт…")
        return await generate_report(chat_id, data)
    c = criteria[step]
    text = (
        f"<b>Вопрос {step+1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )
    kb = InlineKeyboardBuilder()
    start = 0 if c['max']==1 else 1
    for i in range(start, c['max']+1): kb.button(text=str(i), callback_data=f"score_{i}")
    if step>0: kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)
    await bot.send_message(
        chat_id, text,
        parse_mode=ParseMode.HTML,
        reply_markup=kb.as_markup()
    )

async def generate_report(chat_id: int, data):
    name = data["name"]; ts = data["start"]; ph = data.get("pharmacy","")
    wb = load_workbook(TEMPLATE_PATH); ws = wb.active
    title = (
        f"Отчёт по проверке аптеки\n"
        f"Исполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    )
    ws.merge_cells("A1:G2"); ws["A1"] = title; ws["A1"].font = Font(size=14,bold=True)
    ws["B3"] = ph
    headers = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата"]
    for idx,h in enumerate(headers,1): ws.cell(row=5,column=idx,value=h).font = Font(bold=True)
    row,tots,totm = 6,0,0
    for itm in data["data"]:
        c=itm["crit"]; sc=itm["score"]
        ws.cell(row,1,c['block']); ws.cell(row,2,c['criterion']); ws.cell(row,3,c['requirement'])
        ws.cell(row,4,sc); ws.cell(row,5,c['max']); ws.cell(row,7,ts)
        tots+=sc; totm+=c['max']; row+=1
    ws.cell(row+1,3,"ИТОГО:"); ws.cell(row+1,4,tots)
    ws.cell(row+2,3,"Максимум:"); ws.cell(row+2,4,totm)
    fn = f"{ph}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx".replace(" ","_")
    wb.save(fn)
    # Отправка
    with open(fn,'rb') as f: await bot.send_document(QA_CHAT_ID, FSInputFile(f, filename=fn))
    with open(fn,'rb') as f: await bot.send_document(chat_id, FSInputFile(f, filename=fn))
    os.remove(fn)
    log_csv(ph, name, ts, tots, totm)
    await bot.send_message(chat_id, "✅ Отчёт отправлен. Для новой проверки — /start")

# === Webhook setup ===
async def handle_webhook(request: web.Request):
    data = await request.json(); upd = Update(**data)
    await dp.feed_update(bot=bot, update=upd)
    return web.Response(text="OK")
async def health(request: web.Request): return web.Response(text="OK")
app = web.Application()
app.router.add_get("/", health)
app.router.add_post("/webhook", handle_webhook)
app.on_startup.append(lambda app: bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True))
app.on_cleanup.append(lambda app: bot.delete_webhook())

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    web.run_app(app, host="0.0.0.0", port=PORT)
```

**Что изменено:**
- Добавлена опция `0` при `max=1`.
- Адаптация под Render (Web Service + webhook).
- Финальное сообщение с `/start` для повторного запуска.
- Отправка отчёта и в QA-чат (`CHAT_ID`), и пользователю.
