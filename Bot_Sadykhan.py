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

# === ЗАГРУЗКА ОКРУЖЕНИЯ ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")      # https://<your-app>.onrender.com/webhook
PORT           = int(os.getenv("PORT", "8000"))

ALLOWED_USERS = ["Николай Крылов","Таждин Усейн","Жанар Бөлтірік","Шара Абдиева","Тохтар Чарабасов","*"]

# === FSM ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Загрузка критериев ===
criteria_df = pd.read_excel(CHECKLIST_PATH, sheet_name='Чек лист', header=None)
start_i = criteria_df[criteria_df.iloc[:,0]=="Блок"].index[0] + 1
criteria_df = criteria_df.iloc[start_i:,:8].reset_index(drop=True)
criteria_df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
criteria_df = criteria_df.dropna(subset=["Критерий","Требование"])
criteria_list, last_block = [], None
for _, r in criteria_df.iterrows():
    block = r["Блок"] if pd.notna(r["Блок"]) else last_block
    last_block = block
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria_list.append({"block": block, "criterion": r["Критерий"], "requirement": r["Требование"], "max": maxv})

# === Утилиты ===
def get_astana_time():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_submission(pharmacy, name, ts, score, max_score):
    exists = os.path.exists(LOG_PATH)
    with open(LOG_PATH, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["Дата","Аптека","ФИО","Баллы","Макс"])
        w.writerow([ts, pharmacy, name, score, max_score])

# === Инициализация ===
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)

# === Хендлеры команд ===
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("📋 <b>Чек-лист посещения аптек</b>\nВведите ФИО:", parse_mode=ParseMode.HTML)
    await state.set_state(Form.name)

@dp.message(F.text == "/id")
async def cmd_id(msg: types.Message):
    await msg.answer(f"Ваш chat_id: <code>{msg.chat.id}</code>", parse_mode=ParseMode.HTML)

@dp.message(F.text == "/лог")
async def cmd_log(msg: types.Message):
    if os.path.exists(LOG_PATH):
        await msg.answer_document(FSInputFile(LOG_PATH))
    else:
        await msg.answer("Лог ещё не сформирован.")

@dp.message(F.text == "/сброс")
async def cmd_reset(msg: types.Message, state: FSMContext):
    await state.clear()
    await msg.answer("Состояние сброшено. /start — чтобы начать заново")

# === Авторизация ===
@dp.message(Form.name)
async def process_name(msg: types.Message, state: FSMContext):
    user = msg.text.strip()
    if user in ALLOWED_USERS or "*" in ALLOWED_USERS:
        await state.update_data(name=user, step=0, data=[], start=get_astana_time())
        await msg.answer("Введите название аптеки:")
        await state.set_state(Form.pharmacy)
    else:
        await msg.answer("ФИО не распознано. Обратитесь в ИТ-отдел.")

@dp.message(Form.pharmacy)
async def process_pharmacy(msg: types.Message, state: FSMContext):
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_criterion(msg.chat.id, state)

# === Обработка оценок ===
@dp.callback_query(F.data.startswith("score_") | F.data == "prev")
async def score_handler(cb: types.CallbackQuery, state: FSMContext):
    logging.info(f"Raw callback data: {cb.data}")
    await cb.answer("✔️")
    data = await state.get_data(); step = data.get("step",0)
    if cb.data=="prev" and step>0:
        data["step"]-=1; data["data"].pop()
        await state.set_data(data)
        return await send_criterion(cb.from_user.id, state)
    sc = int(cb.data.split("_")[1])
    if step < len(criteria_list):
        data.setdefault("data",[]).append({"crit":criteria_list[step],"score":sc})
        data["step"]+=1
        await state.set_data(data)
    await bot.edit_message_text(cb.message.chat.id, cb.message.message_id,
        f"✅ Оценка: {sc} {'⭐'*sc}")
    await send_criterion(cb.from_user.id, state)

async def send_criterion(chat_id, state: FSMContext):
    data = await state.get_data(); step=data["step"]; total=len(criteria_list)
    if step>=total:
        await bot.send_message(chat_id,"Проверка завершена.формируем отчёт…")
        await generate_and_send(chat_id,data)
        await bot.send_message(chat_id,"Готово! /start")
        return await state.clear()
    c=criteria_list[step]
    msg=(f"<b>Вопрос {step+1} из {total}</b>\n\n"
         f"<b>Блок:</b> {c['block']}\n\n"
         f"<b>Критерий:</b> {c['criterion']}\n\n"
         f"<b>Требование:</b> {c['requirement']}\n\n"
         f"<b>Макс. балл:</b> {c['max']}")
    kb=InlineKeyboardBuilder()
    start=0 if c["max"]==1 else 1
    for i in range(start,c["max"]+1):
        kb.button(text=str(i),callback_data=f"score_{i}")
    if step>0: kb.button(text="◀️ Назад",callback_data="prev")
    kb.adjust(5)
    await bot.send_message(chat_id,msg,parse_mode=ParseMode.HTML,reply_markup=kb.as_markup())

async def generate_and_send(chat_id,data):
    name=data["name"]; ts=data["start"]; pharm=data.get("pharmacy","Без названия")
    wb=load_workbook(TEMPLATE_PATH); ws=wb.active
    title=(f"Отчёт по проверке аптеки\nИсполнитель: {name}\n"
           f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}")
    ws.merge_cells("A1:G2"); ws["A1"]=title; ws["A1"].font=Font(size=14,bold=True)
    ws["B3"]=pharm
    hdr=["Блок","Критерий","Требование","Оценка участника","Макс. оценка","Примечание","Дата проверки"]
    for idx,h in enumerate(hdr,1): cell=ws.cell(row=5,column=idx,value=h); cell.font=Font(bold=True)
    row=6; tscr=0; tmax=0
    for it in data["data"]:
        c=it["crit"]; sc=it["score"]
        ws.cell(row,1,c["block"]); ws.cell(row,2,c["criterion"])
        ws.cell(row,3,c["requirement"]); ws.cell(row,4,sc)
        ws.cell(row,5,c["max"]); ws.cell(row,7,ts)
        tscr+=sc; tmax+=c["max"]; row+=1
    ws.cell(row+1,3,"ИТОГО:"); ws.cell(row+1,4,tscr)
    ws.cell(row+2,3,"Максимум:"); ws.cell(row+2,4,tmax)
    fn=f"{pharm}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}.xlsx".replace(" ","_")
    wb.save(fn)
    with open(fn,"rb") as f: await bot.send_document(CHAT_ID,FSInputFile(f))
    os.remove(fn); log_submission(pharm,name,ts,tscr,tmax)

# === Webhook & healthcheck ===
async def handle_webhook(request:web.Request):
    data=await request.json()
    logging.info(f"Raw update: {data}")
    upd=Update(**data)
    await dp.feed_update(bot=bot, update=upd)
    return web.Response(text="OK")

async def health(request:web.Request):
    return web.Response(text="OK")

async def on_startup(app):
    await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True, allowed_updates=[])

async def on_cleanup(app):
    await bot.delete_webhook()
    await storage.close()

app=web.Application()
app.router.add_get("/",health)
app.router.add_post("/webhook",handle_webhook)
app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__=="__main__":
    logging.basicConfig(level=logging.INFO)
    web.run_app(app,host="0.0.0.0",port=PORT)
