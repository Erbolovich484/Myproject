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

# === Настройки ===
load_dotenv()
API_TOKEN      = os.getenv("API_TOKEN")
CHAT_ID        = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH  = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "Упрощенный чек-лист для проверки аптек.xlsx")
LOG_PATH       = os.getenv("LOG_PATH", "checklist_log.csv")
WEBHOOK_URL    = os.getenv("WEBHOOK_URL")    # https://<app>.onrender.com/webhook
PORT           = int(os.getenv("PORT", "8000"))

ALLOWED_USERS = [
    "Николай Крылов", "Таждин Усейн", "Жанар Бөлтірік",
    "Шара Абдиева", "Тохтар Чарабасов", "*"
]

# === FSM-состояния ===
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# === Чтение чек-листа ===
df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = df[df.iloc[:,0]=="Блок"].index[0] + 1
df = df.iloc[start_i:,:8].reset_index(drop=True)
df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение","Примечание","Дата проверки","Дата исправления"]
df = df.dropna(subset=["Критерий","Требование"])

criteria = []
last = None
for _, r in df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else last
    last = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({"block":blk,"criterion":r["Критерий"],"requirement":r["Требование"],"max":maxv})

# === Утилиты ===
def now_ts():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(ph, nm, ts, sc, mx):
    ex = os.path.exists(LOG_PATH)
    with open(LOG_PATH,"a",newline="",encoding="utf-8") as f:
        w = csv.writer(f)
        if not ex: w.writerow(["Дата","Аптека","ФИО","Баллы","Макс"])
        w.writerow([ts,ph,nm,sc,mx])

# === Инициализация ===
session = AiohttpSession()
bot = Bot(token=API_TOKEN, session=session, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)

# --- Старт и команды ---
@dp.message(F.text=="/start")
async def cmd_start(m: types.Message, s: FSMContext):
    await s.clear()
    await m.answer("📋 <b>Чек-лист посещения аптек</b>\n\nВведите ФИО:", parse_mode=ParseMode.HTML)
    await s.set_state(Form.name)

@dp.message(F.text=="/id")
async def cmd_id(m: types.Message):
    await m.answer(f"Ваш chat_id: <code>{m.chat.id}</code>", parse_mode=ParseMode.HTML)

@dp.message(F.text=="/лог")
async def cmd_log(m: types.Message):
    if os.path.exists(LOG_PATH):
        await m.answer_document(FSInputFile(LOG_PATH))
    else:
        await m.answer("Лог ещё не сформирован.")

@dp.message(F.text=="/сброс")
async def cmd_reset(m: types.Message, s: FSMContext):
    await s.clear()
    await m.answer("Сброшено. /start")

# --- Авторизация ---
@dp.message(Form.name)
async def proc_name(m: types.Message, s: FSMContext):
    user = m.text.strip()
    if user in ALLOWED_USERS or "*" in ALLOWED_USERS:
        await s.update_data(name=user, step=0, data=[], start=now_ts())
        await m.answer("Введите название аптеки:")
        await s.set_state(Form.pharmacy)
    else:
        await m.answer("ФИО не распознано.")

@dp.message(Form.pharmacy)
async def proc_pharmacy(m: types.Message, s: FSMContext):
    await s.update_data(pharmacy=m.text.strip())
    await m.answer("Начинаем проверку…")
    await s.set_state(Form.rating)
    await send_q(m.chat.id, s)

# --- CallbackQuery с фильтром и немедленным ответом ---
@dp.callback_query(F.data.startswith("score_")|F.data=="prev")
async def cb_handler(cb: types.CallbackQuery, s: FSMContext):
    # 1) сразу ответить Telegram
    await cb.answer()

    data = await s.get_data()
    step = data.get("step",0)

    # 2) кнопка Назад
    if cb.data=="prev" and step>0:
        data["step"]-=1; data["data"].pop()
        await s.set_data(data)
        return await send_q(cb.from_user.id, s)

    # 3) оценка
    sc = int(cb.data.split("_")[1])
    if step<len(criteria):
        data.setdefault("data",[]).append({"crit":criteria[step],"score":sc})
        data["step"]+=1; await s.set_data(data)

    # 4) редактируем сообщение
    await bot.edit_message_text(cb.message.chat.id,cb.message.message_id,f"✅ Оценка: {sc} {'⭐'*sc}")

    # 5) следующий вопрос
    return await send_q(cb.from_user.id, s)

# --- Отправка вопроса ---
async def send_q(uid:int, s: FSMContext):
    data = await s.get_data(); step=data["step"]; total=len(criteria)
    if step>=total:
        await bot.send_message(uid,"Проверка завершена. Формируем отчёт…")
        return await gen_report(uid,data)

    c=criteria[step]
    txt = (f"<b>Вопрос {step+1}/{total}</b>\n\n"
           f"<b>Блок:</b> {c['block']}\n"
           f"<b>Критерий:</b> {c['criterion']}\n"
           f"<b>Требование:</b> {c['requirement']}\n"
           f"Макс. балл: {c['max']}")
    kb=InlineKeyboardBuilder()
    start = 0 if c["max"]==1 else 1
    for i in range(start,c["max"]+1): kb.button(str(i),f"score_{i}")
    if step>0: kb.button("◀️ Назад","prev")
    kb.adjust(5)

    await bot.send_message(uid,txt,parse_mode=ParseMode.HTML,reply_markup=kb.as_markup())

# --- Формирование отчёта ---
async def gen_report(uid:int,data):
    name=data["name"]; ts=data["start"]; ph=data.get("pharmacy","")
    wb=load_workbook(TEMPLATE_PATH); ws=wb.active
    title=f"Отчёт по проверке аптеки\nИсполнитель: {name}\nДата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y')}"
    ws.merge_cells("A1:G2"); ws["A1"]=title; ws["A1"].font=Font(size=14,bold=True)
    ws["B3"]=ph
    hdr=["Блок","Критерий","Требование","Оценка","Макс","Примеч.","Дата"]
    for i,h in enumerate(hdr,1): c=ws.cell(row=5,col=i,value=h); c.font=Font(bold=True)
    row=6; tscr=0; tmx=0
    for it in data["data"]:
        cinfo=it["crit"]; sc=it["score"]
        ws.cell(row,1,cinfo["block"]); ws.cell(row,2,cinfo["criterion"])
        ws.cell(row,3,cinfo["requirement"]); ws.cell(row,4,sc)
        ws.cell(row,5,cinfo["max"]); ws.cell(row,7,ts)
        tscr+=sc; tmx+=cinfo["max"]; row+=1
    ws.cell(row+1,3,"ИТОГО:"); ws.cell(row+1,4,tscr)
    ws.cell(row+2,3,"Максимум:"); ws.cell(row+2,4,tmx)
    fn=f"{ph}_{name}_{datetime.strptime(ts,'%Y-%m-%d %H:%M:%S').strftime('%d%m%Y')}.xlsx".replace(" ","_")
    wb.save(fn)
    with open(fn,"rb") as f: await bot.send_document(CHAT_ID,FSInputFile(f,fn))
    with open(fn,"rb") as f: await bot.send_document(uid,FSInputFile(f,fn))
    os.remove(fn); log_csv(ph,name,ts,tscr,tmx)
    await bot.send_message(uid,"✅ Отчёт готов. /start")

# --- Webhook & healthcheck ---
async def handle_webhook(r:web.Request):
    data=await r.json(); upd=Update(**data)
    await dp.feed_update(bot=bot,update=upd)
    return web.Response(text="OK")

async def health(r:web.Request): return web.Response(text="OK")

app=web.Application()
app.router.add_get("/",health)
app.router.add_post("/webhook",handle_webhook)
app.on_startup.append(lambda a: bot.set_webhook(WEBHOOK_URL,drop_pending_updates=True,allowed_updates=["message","callback_query"]))
app.on_cleanup.append(lambda a: bot.delete_webhook())

if __name__=="__main__":
    logging.basicConfig(level=logging.INFO)
    web.run_app(app,host="0.0.0.0",port=PORT)
