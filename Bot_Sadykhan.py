import logging
import os
import csv
import pytz
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import asyncio

from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ParseMode
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import FSInputFile, Update
from aiohttp import web

# === Настройка логирования ===
logging.basicConfig(level=logging.INFO)

# === Загрузка конфигов ===
load_dotenv()
API_TOKEN = os.getenv("API_TOKEN")
CHAT_ID = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_PATH", "checklist.xlsx")  # Исправлено имя файла по умолчанию
LOG_PATH = os.getenv("LOG_PATH", "checklist_log.csv")
PORT = int(os.getenv("PORT", 8080))  # Убедитесь, что этот порт соответствует настройкам Railway
WEBHOOK_URL = os.getenv("WEBHOOK_URL")

# === FSM-состояния ===
class Form(StatesGroup):
    name = State()
    pharmacy = State()
    rating = State()
    comment = State()

# === Читаем критерии из Excel ===
try:
    df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
    start_i = df[df.iloc[:, 0] == "Блок"].index[0] + 1
    df = df.iloc[start_i:, :8].dropna(subset=[1, 2]).reset_index(drop=True)
    df.columns = ["Блок", "Критерий", "Требование", "Оценка", "Макс", "Примечание", "Дата проверки",
                  "Дата исправления"]

    criteria = []
    last = None
    for _, r in df.iterrows():
        blk = r["Блок"] if pd.notna(r["Блок"]) else last
        last = blk
        maxv = int(r["Макс"]) if str(r["Макс"]).isdigit() else 10
        criteria.append({
            "block": blk,
            "criterion": r["Критерий"],
            "requirement": r["Требование"],
            "max": maxv
        })
except FileNotFoundError:
    logging.error(f"Файл не найден: {CHECKLIST_PATH}")
    criteria = []
except Exception as e:
    logging.error(f"Ошибка при чтении файла {CHECKLIST_PATH}: {e}")
    criteria = []

# === Утилиты ===
def now_ts():
    return datetime.now(pytz.timezone("Asia/Almaty")).strftime("%Y-%m-%d %H:%M:%S")

def log_csv(pharm, name, ts, score, total):
    first = not os.path.exists(LOG_PATH)
    try:
        with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if first:
                w.writerow(["Дата", "Аптека", "Проверяющий", "Баллы", "Макс"])
            w.writerow([ts, pharm, name, score, total])
    except Exception as e:
        logging.error(f"Ошибка при записи в лог-файл {LOG_PATH}: {e}")

# === Инициализация бота ===
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML)
dp = Dispatcher(storage=MemoryStorage())

# === Команда /start ===
@dp.message(F.text=="/start")
async def cmd_start(msg: types.Message, state: FSMContext):
    logging.info(f"Handler 'cmd_start' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    await state.clear()
    await msg.answer(
        "<b>📋 Чек‑лист посещения аптек</b>\n\n"
        "Этот бот — интеллектуальная собственность ИТ «Садыхан».  \n"
        "Заполняйте чек‑лист вдумчиво и внимательно:  \n"
        "- inline‑кнопки для быстрой оценки;  \n"
        "- если оценка займёт больше минуты — после всех баллов вы сможете написать вывод ручкой.\n\n"
        "Введите ваше ФИО для авторизации:",
        parse_mode=ParseMode.HTML
    )
    await state.set_state(Form.name)
    logging.info(f"User {msg.from_user.id} set state to Form.name")

# === /id для отладки ===
@dp.message(F.text=="/id")
async def cmd_id(msg: types.Message):
    logging.info(f"Handler 'cmd_id' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    await msg.answer(f"<code>{msg.chat.id}</code>")

# === Сброс FSM ===
@dp.message(F.text=="/сброс")
async def cmd_reset(msg: types.Message, state: FSMContext):
    logging.info(f"Handler 'cmd_reset' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново")
    logging.info(f"User {msg.from_user.id} state cleared")

# === Обработка ФИО ===
@dp.message(Form.name)
async def proc_name(msg: types.Message, state: FSMContext):
    logging.info(f"Handler 'proc_name' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    name = msg.text.strip()
    await state.update_data(name=name, step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)
    logging.info(f"User {msg.from_user.id} entered name: {name}, set state to Form.pharmacy, state data: {await state.get_data()}")

# === Обработка названия аптеки ===
@dp.message(Form.pharmacy)
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    logging.info(f"Handler 'proc_pharmacy' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    await state.update_data(pharmacy=msg.text.strip())
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    logging.info("Calling send_question from proc_pharmacy")
    await send_question(msg.chat.id, state)
    logging.info(f"User {msg.from_user.id} entered pharmacy: {await state.get_data()}, set state to Form.rating")

# === Общий хэндлер callback_query ===
@dp.callback_query()
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    logging.info(f"*** CALLBACK QUERY RECEIVED: {cb.data} ***")
    logging.info(f"Callback query received from user {cb.from_user.id}, chat {cb.message.chat.id}, data: {cb.data}")
    data = await state.get_data()
    step = data.get("step", 0)
    total = len(criteria)
    # если уже все оценили — просто acknowledge
    if step >= total:
        return await cb.answer()
    # парсим оценку
    if cb.data.startswith("score_"):
        score = int(cb.data.split("_", 1)[1])
        record = {"crit": criteria[step], "score": score}
        data["data"].append(record)
        data["step"] += 1
        await state.update_data(**data)
        # редактируем исходное сообщение
        await bot.edit_message_text(
            f"✅ Оценка: {score} {'⭐' * score}",
            chat_id=cb.message.chat.id,
            message_id=cb.message.message_id
        )
        # отправляем следующий
        return await send_question(cb.from_user.id, state)
    # навигация «Назад»
    if cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.update_data(**data)
        return await send_question(cb.from_user.id, state)

# === Функция отправки следующего вопроса или финального промпта ===
async def send_question(chat_id: int, state: FSMContext):
    data = await state.get_data()
    step = data.get("step", 0)
    total = len(criteria)
    logging.info(f"send_question called. step: {step}, total: {total}, criteria length: {len(criteria)}")

    # если всё — переходим к комментарию
    if step >= total:
        await bot.send_message(
            chat_id,
            "✅ Все оценки поставлены!\n\n"
            "📝 Теперь напишите, пожалуйста, ваши выводы по аптеке:",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(Form.comment)
        logging.info(f"User {chat_id} finished rating, set state to Form.comment")
        return

    c = criteria[step]
    text = (
        f"<b>Вопрос {step + 1} из {total}</b>\n\n"
        f"<b>Блок:</b> {c['block']}\n"
        f"<b>Критерий:</b> {c['criterion']}\n"
        f"<b>Требование:</b> {c['requirement']}\n"
        f"<b>Макс. балл:</b> {c['max']}"
    )
    kb = InlineKeyboardBuilder()
    start = 0 if c["max"] == 1 else 1
    for i in range(start, c["max"] + 1):
        kb.button(text=str(i), callback_data=f"score_{i}")
    if step > 0:
        kb.button(text="◀️ Назад", callback_data="prev")
    kb.adjust(5)

    await bot.send_message(chat_id, text, reply_markup=kb.as_markup())

# === Обработка текстового комментария ===
@dp.message(Form.comment)
async def proc_comment(msg: types.Message, state: FSMContext):
    logging.info(f"Handler 'proc_comment' called by user {msg.from_user.id}, chat {msg.chat.id}, text: {msg.text}")
    data = await state.get_data()
    data["comment"] = msg.text.strip()
    await state.update_data(**data)
    await msg.answer("⌛ Формирую отчёт…")
    await make_report(msg.chat.id, data)
    await state.clear()
    logging.info(f"User {msg.from_user.id} entered comment, report initiated, state cleared")

# === Генерация и отправка отчёта ===
async def make_report(user_id: int, data):
    logging.info(f"Generating report for user {user_id}, data: {data}")
    name = data["name"]
    ts = data["start"]
    pharmacy = data["pharmacy"]
    report_filename = f"{pharmacy}_{name}_{datetime.strptime(ts, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y_%H%M')}.xlsx".replace(" ", "_")

    try:
        logging.info(f"Attempting to load template: {TEMPLATE_PATH}")
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        # Заголовок
        title = (f"Отчёт по проверке аптеки\n"
                 f"Исполнитель: {name}\n"
                 f"Дата и время: {ts}")
        ws.merge_cells("A1:G2")
        ws["A1"] = title
        ws["A1"].font = Font(size=14, bold=True)
        ws["B3"] = pharmacy

        # Шапка таблицы
        headers = ["Блок", "Критерий", "Требование", "Оценка", "Макс", "Коммент.", "Дата проверки"]
        for idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, idx, h)
            cell.font = Font(bold=True)

        # Заполняем строки
        row = 6
        total_score = 0
        total_max = 0
        for rec in data["data"]:
            c = rec["crit"]
            sc = rec["score"]
            ws.cell(row, 1, c["block"])
            ws.cell(row, 2, c["criterion"])
            ws.cell(row, 3, c["requirement"])
            ws.cell(row, 4, sc)
            ws.cell(row, 5, c["max"])
            ws.cell(row, 6, "")  # можно сюда вставить индивидуальный прим.
            ws.cell(row, 7, ts)
            total_score += sc
            total_max += c["max"]
            row += 1

        # Итого
        ws.cell(row + 1, 3, "ИТОГО:")
        ws.cell(row + 1, 4, total_score)
        ws.cell(row + 2, 3, "Максимум:")
        ws.cell(row + 2, 4, total_max)

        # Ваш комментарий внизу
        ws.cell(row + 4, 1, "Вывод проверяющего:")
        ws.cell(row + 5, 1, data.get("comment", ""))

        try:
            logging.info(f"Attempting to save report: {report_filename}")
            wb.save(report_filename)
            logging.info(f"Report '{report_filename}' saved")
        except Exception as e:
            logging.error(f"Error saving report: {e}", exc_info=True)
            await bot.send_message(user_id, "❌ Ошибка при сохранении отчёта.")
            return

        try:
            logging.info(f"Attempting to open report for sending: {report_filename}")
            with open(report_filename, "rb") as f:
                logging.info(f"Report opened successfully, attempting to send to user {user_id}")
                await bot.send_document(user_id, FSInputFile(f, report_filename))
                logging.info(f"Report sent to user {user_id}")
        except Exception as e:
            logging.error(f"Error sending report to user {user_id}: {e}", exc_info=True)
            await bot.send_message(user_id, "❌ Ошибка при отправке отчёта.")
            return

        try:
            logging.info(f"Attempting to open report for sending to chat {CHAT_ID}: {report_filename}")
            with open(report_filename, "rb") as f:
                logging.info(f"Report opened successfully, attempting to send to chat {CHAT_ID}")
                await bot.send_document(CHAT_ID, FSInputFile(f, report_filename))
                logging.info(f"Report sent to chat {CHAT_ID}")
        except Exception as e:
            logging.error(f"Error sending report to chat {CHAT_ID}: {e}", exc_info=True)
            # Не возвращаем здесь, так как отправка в чат не критична для пользователя

        # Логируем
        log_csv(pharmacy, name, ts, total_score, total_max)

        # Финальное сообщение
        await bot.send_message(user_id,
                               "✅ Отчёт сформирован и отправлен.\n"
                               "Для новой проверки — /start")

    except FileNotFoundError:
        logging.error(f"Файл шаблона не найден: {TEMPLATE_PATH}")
        await bot.send_message(user_id, "❌ Ошибка: Файл шаблона отчёта не найден.")
    except Exception as e:
        logging.error(f"Ошибка при создании отчёта: {e}", exc_info=True)
        await bot.send_message(user_id, "❌ Произошла ошибка при формировании отчёта.")
    finally:
        try:
            logging.info(f"Attempting to remove temporary file: {report_filename}")
            os.remove(report_filename)
            logging.info(f"Temporary report file '{report_filename}' deleted")
        except Exception as e:
            logging.warning(f"Не удалось удалить временный файл {report_filename}: {e}")

# === Webhook & запуск ===
async def handle_webhook(request: web.Request):
    logging.info(f"Received webhook request: {request.method} {request.url}")
    try:
        update = Update(**await request.json())
        logging.info(f"Parsed update: {update}")
        await dp.feed_update(bot, update)
        return web.Response(text="OK")
    except Exception as e:
        logging.error(f"Error processing webhook: {e}", exc_info=True)
        return web.Response(status=500)

async def on_startup(bot: Bot):
    if WEBHOOK_URL:
        webhook_url = f"{WEBHOOK_URL}/webhook"
        await bot.set_webhook(webhook_url)
        logging.info(f"Webhook set to: {webhook_url}")
    else:
        logging.warning("WEBHOOK_URL не определен. Бот будет работать в режиме Long Polling.")

async def on_shutdown(bot: Bot):
    logging.warning("Shutting down...")
    await bot.delete_webhook()
    await bot.session.close()
    logging.warning("Bot and session closed.")

async def main():
    dp.startup.register(on_startup)
    dp.shutdown.register(on_shutdown)

    if WEBHOOK_URL:
        app = web.Application()
        app.add_routes([web.post("/webhook", handle_webhook)])
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, "0.0.0.0", PORT)
        await site.start()
        logging.info(f"Web application started on port {PORT}")
        # Keep the server running
        while True:
            await asyncio.sleep(3600)
    else:
        await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
