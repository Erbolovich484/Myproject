import os
import csv
import logging
import pytz
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ParseMode
from aiogram.types import FSInputFile, Update
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.redis import RedisStorage
from aiohttp import web

# ========== Логи ==========
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ========== Конфиг из ENV ==========
API_TOKEN      = os.environ["API_TOKEN"]
CHAT_ID        = int(os.environ["CHAT_ID"])
TEMPLATE_PATH  = os.environ.get("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.environ.get("CHECKLIST_PATH", "checklist.xlsx")
LOG_PATH       = os.environ.get("LOG_PATH", "checklist_log.csv")

# ========== FSM-состояния ==========
class Form(StatesGroup):
    name     = State()
    pharmacy = State()
    rating   = State()

# ========== Чтение критериев ==========
_df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
start_i = _df[_df.iloc[:,0] == "Блок"].index[0] + 1
_df = _df.iloc[start_i:,:8].reset_index(drop=True)
_df.columns = ["Блок","Критерий","Требование","Оценка","Макс. значение",
               "Примечание","Дата проверки","Дата исправления"]
_df = _df.dropna(subset=["Критерий","Требование"])
criteria = []
_last = None
for _, r in _df.iterrows():
    blk = r["Блок"] if pd.notna(r["Блок"]) else _last
    _last = blk
    maxv = int(r["Макс. значение"]) if pd.notna(r["Макс. значение"]) and str(r["Макс. значение"]).isdigit() else 10
    criteria.append({
        "block": blk,
        "criterion": r["Критерий"],
        "requirement": r["Требование"],
        "max": maxv
    })

# ========== Утилиты ==========
def now_str():
    return datetime.now(pytz.timezone("Asia/Almaty"))\
                   .strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharmacy, user, ts, score, mx):
    header = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if header:
            w.writerow(["Дата","Аптека","Исполнитель","Баллы","Макс"])
        w.writerow([ts, pharmacy, user, score, mx])

# ========== Инициализация бота и FSM ==========
storage = RedisStorage.from_url(os.environ["REDIS_URL"])
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp  = Dispatcher(storage=storage)

# ========== /start ==========
@dp.message(F.text == "/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    logger.info("CMD /start from %s", msg.from_user.id)
    await state.clear()
    await msg.answer(
        "📋 <b>Чек-лист посещения аптек</b>\n"
        "Интеллектуальная собственность ИТ-департамента «Садыхан».\n"
        "По завершении отчёт в Excel уйдёт в QA-чат.\n\n"
        "🏁 Введите ваше ФИО:",
    )
    await state.set_state(Form.name)

# ========== Обработка ФИО ==========
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    user = msg.text.strip()
    logger.info("Received name: %s", user)
    # тут можно фильтровать ALLOWED_USERS, если нужно
    await state.update_data(name=user, step=0, data=[], start=now_str())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)

# ========== Обработка названия аптеки ==========
@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    pharm = msg.text.strip()
    logger.info("Received pharmacy: %s", pharm)
    await state.update_data(pharmacy=pharm)
    await msg.answer("Спасибо! Начинаем проверку…")
    await state.set_state(Form.rating)
    await send_question(msg.chat.id, state)

# ========== Callback-кнопки ==========
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    await cb.answer()  # чтобы Telegram не жаловался
    data = await state.get_data()
    step = data["step"]

    # «Назад»
    if cb.data == "prev" and step>0:
        data["step"] -=1
        data["data"].pop()
        await state.set_data(data)
        return await send_question(cb.from_user.id, state)

    # Оценка
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_")[1])
        crit  = criteria[step]
        logger.debug("Callback received: %s", cb.data)
        data.setdefault("data", []).append({"crit": crit, "score": score})
        data["step"] +=1
        await state.set_data(data)

        # редактируем предыдущее сообщение
        await bot.edit_message_text(
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id,
            text=f"✅ Оценка: {score} {'⭐'*score}"
        )
        return await send_question(cb.from_user.id, state)

# ========== Отправка вопроса ==========
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data["step"]
    total= len(criteria)

    if step>=total:
        await bot.send_message(chat_id, "✅ Проверка завершена. Формируем отчёт…")
        return await make_report(chat_id, data)

    c = criteria[step]
    logger.debug("send_question: step=%s/%s to %s", step, total, chat_id)
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
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step>0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# ========== Генерация и отправка отчёта ==========
async def make_report(chat_id: int, data):
    name     = data["name"]
    ts       = data["start"]
    pharmacy = data["pharmacy"]
    entries  = data["data"]

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # заголовок
    ws.merge_cells("A1:G2")
    ws["A1"] = (
        f"Отчёт по проверке аптеки\n"
        f"Исполнитель: {name}\n"
        f"Дата: {datetime.strptime(ts,'%Y-%m-%d %H:%M:%S')\
            .strftime('%d.%m.%Y')}"
    )
    ws["A1"].font = Font(size=14,bold=True)
    ws["B3"] = pharmacy

    # шапка таблицы
    hdr = ["Блок","Критерий","Требование","Оценка","Макс","Примечание","Дата"]
    for i,h in enumerate(hdr,1):
        cell = ws.cell(row=5, column=i, value=h)
        cell.font = Font(bold=True)

    row = 6
    tot_score=0
    tot_max  =0
    for it in entries:
        c  = it["crit"]
        sc = it["score"]
        ws.cell(row,1, c["block"])
        ws.cell(row,2, c["criterion"])
        ws.cell(row,3, c["requirement"])
        ws.cell(row,4, sc)
        ws.cell(row,5, c["max"])
        ws.cell(row,7, ts)
        tot_score += sc
        tot_max   += c["max"]
        row +=1

    # вывод
    ws.cell(row+1,3,"ИТОГО:")
    ws.cell(row+1,4,tot_score)
    ws.cell(row+2,3,"Максимум:")
    ws.cell(row+2,4,tot_max)

    fname = f"{pharmacy}_{name}_{ts[:10]}.xlsx".replace(" ","_")
    wb.save(fname)

    # отправка: в QA-чат и пользователю
    with open(fname,"rb") as f:
        await bot.send_document(CHAT_ID, FSInputFile(f, filename=fname))
    with open(fname,"rb") as f:
        await bot.send_document(chat_id, FSInputFile(f, filename=fname))

    os.remove(fname)
    log_csv(pharmacy, name, ts, tot_score, tot_max)

    await bot.send_message(
        chat_id,
        "📌 Отчёт готов и отправлен в QA-чат.\n"
        "Чтобы пройти снова — /start"
    )

# ========== Webhook & Aiohttp ==========
async def handle_webhook(request: web.Request):
    data = await request.json()
    upd  = Update(**data)
    await dp.feed_update(bot=bot, update=upd)
    return web.Response(text="OK")

app = web.Application()
app.router.add_post("/webhook", handle_webhook)
app.router.add_get("/", lambda r: web.Response(text="OK"))

async def on_startup(app: web.Application):
    logger.info("Setting webhook…")
    await bot.set_webhook(
        f"https://{os.environ['FLY_APP_NAME']}.fly.dev/webhook",
        drop_pending_updates=True
    )

async def on_cleanup(app: web.Application):
    await bot.delete_webhook()
    await storage.close()

app.on_startup.append(on_startup)
app.on_cleanup.append(on_cleanup)

if __name__ == "__main__":
    logging.getLogger("asyncio").setLevel(logging.WARNING)
    web.run_app(app, host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
