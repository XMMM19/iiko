# bot.py
import asyncio
import os
import logging
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ChatMember, FSInputFile, Document
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.exceptions import TelegramBadRequest
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from fileHandler import process_excel

# Загружаем переменные окружения
load_dotenv()

# Настраиваем логгер
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Формат логов
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

# Обработчик для файла
file_handler = logging.FileHandler("bot.log")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# Обработчик для терминала
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# Получаем токен из .env
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN not set in .env file")

# Получаем id группы из .env
CHANNEL_USERNAME = os.getenv("CHANNEL_USERNAME")
if not CHANNEL_USERNAME:
    raise ValueError("CHANNEL_USERNAME not set in .env file")

# --- FSM Состояния ---
class FileProcessing(StatesGroup):
    waiting_for_percentage = State()
    waiting_for_file = State()

# --- Основной хендлер на файл ---
async def handle_document(message: Message, state: FSMContext):
    document: Document = message.document

    if not document.file_name.endswith(".xlsx"):
        await message.answer("Пожалуйста, отправьте файл в формате .xlsx.")
        return

    user_id = message.from_user.id
    if not await check_subscription(message.bot, user_id):
        await message.answer("Пожалуйста, подпишитесь на канал, чтобы пользоваться ботом.")
        return

    file_id = document.file_id
    file_path = f"temp/{document.file_name}"
    await message.bot.download(file_id, destination=file_path)

    await state.update_data(file_path=file_path, file_name=document.file_name)
    await message.answer(
        "Ваш файл принят для дальнейшей обработки.\n\n"
        "Пожалуйста, укажите допустимый процент отклонения. Например:\n"
        "`4.5` или `2`\n\n"
        "Если хотите использовать значение по умолчанию (3.5%), отправьте `Нет` (без точки).",
        parse_mode=ParseMode.MARKDOWN
    )
    await state.set_state(FileProcessing.waiting_for_percentage)

# --- Обработка процента ---
async def handle_percentage(message: Message, state: FSMContext):
    user_input = message.text.strip().lower()
    data = await state.get_data()
    file_path = data["file_path"]
    file_name = data["file_name"]

    try:
        if user_input == "нет":
            percentage = 0.035
        else:
            percentage = float(user_input.replace(",", ".")) / 100.0
    except ValueError:
        await message.answer("Введите число, например `4.5` или `Нет`.")
        return

    output_path = f"temp/processed_{file_name}"
    await message.answer("Идет обработка файла, пожалуйста подождите...")

    try:
        process_excel(file_path, output_path, allowed_deviation_percentage=percentage)
        await message.answer_document(FSInputFile(output_path), caption="Обработка завершена. Вот ваш файл.")
    except Exception as e:
        logging.exception("Ошибка при обработке файла:")
        await message.answer("Произошла ошибка при обработке файла.")

    await state.clear()

async def check_subscription(bot: Bot, user_id: int) -> bool:
    try:
        member: ChatMember = await bot.get_chat_member(chat_id=CHANNEL_USERNAME, user_id=user_id)
        return member.status in ("member", "administrator", "creator")
    except TelegramBadRequest:
        return False  # например, если пользователь не найден в канале

# Обработчик команды /start
async def cmd_start(message: Message):
    logging.info(f"Получена команда /start от пользователя id={message.from_user.id}")
    await message.answer(
        "Здравствуйте!\n"
        "Для активации функций бота, пожалуйста, подпишитесь на наш канал:\n\n"
        f"{CHANNEL_USERNAME}\n\n"
        "После подписки нажмите /check для продолжения."
    )

# Обработчик команды /check
async def check_user_subscription(message: Message, bot: Bot):
    user_id = message.from_user.id
    logging.info(f"Получена команда /check от пользователя id={user_id}")
    if await check_subscription(bot, user_id):
        await message.answer("Вы подписаны на канал. Теперь вы можете отправлять файл на обработку.")
    else:
        await message.answer("Подписка не найдена. Пожалуйста, подпишитесь на канал.")

# Основная функция
async def main():
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    # await bot.delete_webhook(drop_pending_updates=True)

    dp = Dispatcher(storage=MemoryStorage())
    dp.message.register(cmd_start, F.text == "/start")
    dp.message.register(check_user_subscription, F.text == "/check")
    dp.message.register(handle_document, F.document)
    dp.message.register(handle_percentage, FileProcessing.waiting_for_percentage)

    os.makedirs("temp", exist_ok=True)
    logging.info("Бот запущен.")
    await dp.start_polling(bot)

# Запуск
if __name__ == "__main__":
    asyncio.run(main())
