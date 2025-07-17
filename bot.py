# bot.py
import asyncio
import os
import logging
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ChatMember
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.exceptions import TelegramBadRequest

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
        await message.answer("Вы подписаны на канал.")
    else:
        await message.answer("Подписка не найдена. Пожалуйста, подпишитесь на канал.")

# Основная функция
async def main():
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    # await bot.delete_webhook(drop_pending_updates=True)

    dp = Dispatcher(storage=MemoryStorage())
    dp.message.register(cmd_start, F.text == "/start")
    dp.message.register(check_user_subscription, F.text == "/check")

    logging.info("Бот запущен.")
    await dp.start_polling(bot)

# Запуск
if __name__ == "__main__":
    asyncio.run(main())
