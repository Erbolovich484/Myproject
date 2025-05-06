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
from aiogram.client.default import DefaultBotProperties

# === Настройка логирования ===
logging.basicConfig(level=logging.DEBUG)

# === Загрузка конфигов ===
load_dotenv()
API_TOKEN = os.getenv("API_TOKEN")
CHAT_ID = int(os.getenv("CHAT_ID", "0"))
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "template.xlsx")
CHECKLIST_PATH = os.getenv("CHECKLIST_DATA", "checklist.xlsx")
LOG_PATH = os.getenv("LOG_PATH", "checklist_log.csv")
PORT = int(os.getenv("PORT", 8080))
WEBHOOK_URL = os.getenv("WEBHOOK_URL")

logging.info("Bot configuration loaded.")
logging.debug(f"API_TOKEN is set: {API_TOKEN is not None}")
logging.debug(f"CHAT_ID: {CHAT_ID}")
logging.debug(f"TEMPLATE_PATH: {TEMPLATE_PATH}")
logging.debug(f"CHECKLIST_PATH: {CHECKLIST_PATH}")
logging.debug(f"LOG_PATH: {LOG_PATH}")
logging.debug(f"PORT: {PORT}")
logging.debug(f"WEBHOOK_URL: {WEBHOOK_URL}")

# === FSM-состояния ===
class Form(StatesGroup):
    name = State()
    pharmacy = State()
    rating = State()
    comment = State()

# === Читаем критерии из Excel ===
criteria = []
try:
    logging.info(f"Attempting to read checklist from: {CHECKLIST_PATH}")
    df = pd.read_excel(CHECKLIST_PATH, sheet_name="Чек лист", header=None)
    logging.debug(f"Checklist DataFrame shape: {df.shape}")
    start_i = df[df.iloc[:, 0] == "Блок"].index[0] + 1
    logging.debug(f"Start index for data: {start_i}")
    df = df.iloc[start_i:, :8].dropna(subset=[1, 2]).reset_index(drop=True)
    logging.debug(f"Filtered DataFrame shape: {df.shape}")
    df.columns = ["Блок", "Критерий", "Требование", "Оценка", "Макс", "Примечание", "Дата проверки",
                  "Дата исправления"]
    logging.debug(f"DataFrame columns: {df.columns.tolist()}")

    last = None
    for _, r in df.iterrows():
        logging.debug(f"Processing row:\n{r}")
        blk = r["Блок"] if pd.notna(r["Блок"]) else last
        last = blk
        maxv_str = str(r["Макс"])
        if maxv_str.isdigit():
            maxv = int(maxv_str)
        else:
            logging.warning(f"Invalid value in 'Макс' column: '{maxv_str}'. Using default 10.")
            maxv = 10
        criteria.append({
            "block": blk,
            "criterion": r["Критерий"],
            "requirement": r["Требование"],
            "max": maxv
        })
    logging.info(f"Successfully loaded {len(criteria)} criteria.")
except FileNotFoundError as e:
    logging.error(f"Checklist file not found at: {CHECKLIST_PATH} - {e}")
except ValueError as e:
    logging.error(f"Error processing Excel file (ValueError): {e}. Check file format and content.")
except KeyError as e:
    logging.error(f"Error accessing column in Excel: {e}. Check column names.")
except Exception as e:
    logging.error(f"An unexpected error occurred while reading the checklist: {e}", exc_info=True)

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
        logging.error(f"Error writing to log file: {LOG_PATH} - {e}", exc_info=True)

# === Инициализация бота ===
bot = Bot(token=API_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher(storage=MemoryStorage())

# === Команда /start ===
async def cmd_start(msg: types.Message, state: FSMContext):
    logging.info(f"User {msg.from_user.id} started the bot.")
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
    logging.debug(f"User {msg.from_user.id} entered state: Form.name")

# === /id для отладки ===
async def cmd_id(msg: types.Message):
    logging.info(f"User {msg.from_user.id} requested their chat ID.")
    await msg.answer(f"<code>{msg.chat.id}</code>")

# === Сброс FSM ===
async def cmd_reset(msg: types.Message, state: FSMContext):
    logging.info(f"User {msg.from_user.id} requested state reset.")
    await state.clear()
    await msg.answer("Состояние сброшено. /start — начать заново")
    logging.debug(f"User {msg.from_user.id} state cleared.")

# === Обработка ФИО ===
async def proc_name(msg: types.Message, state: FSMContext):
    name = msg.text.strip()
    logging.info(f"User {msg.from_user.id} entered name: {name}")
    await state.update_data(name=name, step=0, data=[], start=now_ts())
    await msg.answer("Введите название аптеки:")
    await state.set_state(Form.pharmacy)
    logging.debug(f"User {msg.from_user.id} entered state: Form.pharmacy, data: {await state.get_data()}")

# === Обработка названия аптеки ===
async def proc_pharmacy(msg: types.Message, state: FSMContext):
    logging.info(f"proc_pharmacy started for user {msg.from_user.id}")
    pharmacy = msg.text.strip()
    logging.info(f"User {msg.from_user.id} entered pharmacy: {pharmacy}")
    await state.update_data(pharmacy=pharmacy)
    await msg.answer("Начинаем проверку…")
    await state.set_state(Form.rating)
    logging.debug(f"User {msg.from_user.id} entered state: Form.rating, data: {await state.get_data()}")
    logging.info(f"Calling send_question from proc_pharmacy for user {msg.from_user.id}")
    await send_question(msg.chat.id, state)
    logging.info(f"proc_pharmacy finished for user {msg.from_user.id}")

# === Общий хэндлер callback_query ===
async def cb_all(cb: types.CallbackQuery, state: FSMContext):
    logging.info(f"Callback query received from user {cb.from_user.id}, data: {cb.data}")
    data = await state.get_data()
    step = data.get("step", 0)
    total = len(criteria)
    logging.debug(f"Current step: {step}, Total criteria: {total}, Callback data: {cb.data}")

    if step >= total:
        logging.debug("All criteria have been rated.")
        return await cb.answer()

    if cb.data.startswith("score_"):
        score = int(cb.data.split("_", 1)[1])
        record = {"crit": criteria[step], "score": score}
        data["data"].append(record)
        data["step"] += 1
        await state.update_data(**data)
        try:
            await bot.edit_message_text(
                f"✅ Оценка: {score} {'⭐' * score}",
                chat_id=cb.message.chat.id,
                message_id=cb.message.message_id
            )
        except Exception as e:
            logging.error(f"Error editing message in cb_all: {e}", exc_info=True)
        await send_question(cb.from_user.id, state)
        return
    elif cb.data == "prev" and step > 0:
        data["step"] -= 1
        data["data"].pop()
        await state.update_data(**data)
        logging.debug(f"User {cb.from_user.id} navigated back to criterion {data['step'] + 1}")
        await send_question(cb.from_user.id, state)
        return
    else:
        logging.warning(f"Unhandled callback data: {cb.data}")
        await cb.answer("Неизвестная команда")

# === Функция отправки следующего вопроса или финального промпта ===
async def send_question(chat_id: int, state: FSMContext):
    logging.info(f"send_question started for chat {chat_id}")
    data = await state.get_data()
    step = data.get("step", 0)
    total = len(criteria)
    logging.info(f"Sending question {step + 1} of {total} to chat {chat_id}.")
    logging.debug(f"Current state data in send_question: {data}")

    if step >= total:
        logging.info(f"All {total} criteria have been processed for chat {chat_id}.")
        await bot.send_message(
            chat_id,
            "✅ Все оценки поставлены!\n\n"
            "📝 Теперь напишите, пожалуйста, ваши выводы по аптеке:",
            parse_mode=ParseMode.HTML
        )
        await state.set_state(Form.comment)
        logging.debug(f"User {chat_id} entered state: Form.comment")
        return

    if not criteria:
        logging.error("Criteria list is empty. Check checklist file.")
        await bot.send_message(chat_id, "❌ Ошибка: Не удалось загрузить критерии проверки.")
        await state.clear()
        return

    if step < len(criteria):
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

        try:
            await bot.send_message(chat_id, text, reply_markup=kb.as_markup())
            logging.debug(f"Question sent to chat {chat_id} with keyboard: {kb.as_markup()}")
        except Exception as e:
            logging.error(f"Error sending message in send_question: {e}", exc_info=True)
    else:
        logging.warning(f"Attempted to access criterion with index {step}, but total criteria is {total}.")

# === Обработка текстового комментария ===
async def proc_comment(msg: types.Message, state: FSMContext):
    comment = msg.text.strip()
    logging.info(f"User {msg.from_user.id} entered comment: {comment}")
    data = await state.get_data()
    data["comment"] = comment
    await state.update_data(**data)
    await msg.answer("⌛ Формирую отчёт…")
    await make_report(msg.chat.id, data)
    await state.clear()
    logging.debug(f"User {msg.from_user.id} entered comment, report initiated, state cleared.")

# === Генерация и отправка отчёта ===
async def make_report(user_id: int, data):
    logging.info(f"Generating report for user {user_id} with data:\n{data}")
    name = data["name"]
    ts = data["start"]
    pharmacy = data["pharmacy"]
    report_filename = f"{pharmacy}_{name}_{datetime.strptime(ts, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y_%H%M')}.xlsx".replace(" ", "_")

    try:
        logging.info(f"Attempting to load template from: {TEMPLATE_PATH}")
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        title = (f"Отчёт по проверке аптеки\n"
                 f"Исполнитель: {name}\n"
                 f"Дата и время: {ts}")
        ws.merge_cells("A1:G2")
        ws["A1"] = title
        ws["A1"].font = Font(size=14, bold=True)
        ws["B3"] = pharmacy

        headers = ["Блок", "Критерий", "Требование", "Оценка", "Макс", "Коммент.", "Дата проверки"]
        for idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, idx, h)
            cell.font = Font(bold=True)

        row = 6
        total_score = 0
        total_max = 0
        for rec in data.get("data", []):
            c = rec["crit"]
            sc = rec["score"]
            ws.cell(row, 1, c["block"])
            ws.cell(row, 2, c["criterion"])
            ws.cell(row, 3, c["requirement"])
            ws.cell(row, 4, sc)
            ws.cell(row, 5, c["max"])
            ws.cell(row, 6, "")
            ws.cell(row, 7, ts)
            total_score += sc
            total_max += c["max"]
            row += 1

        ws.cell(row + 1, 3, "ИТОГО:")
        ws.cell(row + 1, 4, total_score)
        ws.cell(row + 2, 3, "Максимум:")
        ws.cell(row + 2, 4, total_max)

        ws.cell(row + 4, 1, "Вывод проверяющего:")
        ws.cell(row + 5, 1, data.get("comment", ""))

        try:
            logging.info(f"Attempting to save report to: {report_filename}")
            wb.save(report_filename)
            logging.info(f"Report savedsuccessfully.")
        except Exception as e:
            logging.error(f"Error saving report: {e}", exc_info=True)
            await bot.send_message(user_id, "❌ Ошибка при сохранении отчёта.")
            return

        try:
            logging.info(f"Attempting to send report to user {user_id}.")
            with open(report_filename, "rb") as f:
                await bot.send_document(user_id, FSInputFile(f, filename=report_filename))
            logging.info(f"Report sent to user {user_id}.")
        except Exception as e:
            logging.error(f"Error sending report: {e}", exc_info=True)
            await bot.send_message(user_id, "❌ Ошибка при отправке отчёта.")
        finally:
            try:
                logging.info(f"Attempting to remove temporary file: {report_filename}")
                os.remove(report_filename)
                logging.info(f"Temporary report file '{report_filename}' deleted.")
            except Exception as e:
                logging.warning(f"Failed to delete temporary file {report_filename}: {e}")

        await bot.send_message(user_id,
                                    "✅ Отчёт сформирован и отправлен.\n"
                                    "Для новой проверки — /start")

    except FileNotFoundError:
        logging.error(f"Template file not found at: {TEMPLATE_PATH}")
        await bot.send_message(user_id, "❌ Ошибка: Файл шаблона отчёта не найден.")
    except Exception as e:
        logging.error(f"An unexpected error occurred during report generation: {e}", exc_info=True)
        await bot.send_message(user_id, "❌ Произошла ошибка при формировании отчёта.")
    finally:
        try:
            logging.info(f"Attempting to remove temporary file: {report_filename}")
            os.remove(report_filename)
            logging.info(f"Temporary report file '{report_filename}' deleted.")
        except Exception as e:
            logging.warning(f"Failed to delete temporary file {report_filename}: {e}")

# === Webhook & запуск ===
async def handle_webhook(request: web.Request):
    logging.info(f"Received webhook request: {request.method} {request.url}")
    try:
        update = await request.json()
        logging.info(f"Webhook data: {update}")  # Log the received data
        update = Update(**update)
        await dp.feed_update(bot, update)
        return web.Response(text="OK")
    except Exception as e:
        logging.error(f"Error processing webhook: {e}", exc_info=True)
        return web.Response(status=500)

async def on_startup(bot: Bot):
    if WEBHOOK_URL:
        webhook_url = f"{WEBHOOK_URL}/webhook"
        try:
            await bot.set_webhook(webhook_url)
            webhook_info = await bot.get_webhook_info()
            logging.info(f"Webhook set to: {webhook_url}")
            logging.info(f"Current webhook status: {webhook_info}")
            if webhook_info.last_error_date:
                logging.error(f"Last webhook error: {webhook_info.last_error_date} - {webhook_info.last_error_message}")
        except Exception as e:
            logging.error(f"Error setting webhook: {e}", exc_info=True)
    else:
        logging.warning("WEBHOOK_URL is not defined. Bot will run in Long Polling mode.")

async def on_shutdown(bot: Bot):
    logging.warning("Shutting down bot...")
    await bot.delete_webhook()
    await bot.session.close()
    logging.info("Bot session closed.")

async def main():
    dp.startup.register(on_startup)
    dp.shutdown.register(on_shutdown)

    # Регистрация обработчиков
    dp.message.register(cmd_start, F.text=="/start")
    dp.message.register(cmd_id, F.text=="/id")
    dp.message.register(cmd_reset, F.text=="/сброс")
    dp.message.register(proc_name, Form.name)
    dp.message.register(proc_pharmacy, Form.pharmacy)
    dp.message.register(proc_comment, Form.comment)
    dp.callback_query.register(cb_all)

    if WEBHOOK_URL:
        app = web.Application()
        app.add_routes([web.post("/webhook", handle_webhook)])
        runner = web.AppRunner(app)
        await runner.setup()
        site = web.TCPSite(runner, "0.0.0.0", PORT)
        await site.start()
        logging.info(f"Web application started on port {PORT}")
        # Бесконечный цикл для работы веб-сервера
        while True:
            await asyncio.sleep(3600)
    else:
        await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
