# ============================================
#  АВТОУСТАНОВКА БИБЛИОТЕК
# ============================================

import subprocess
import sys
import os

def install_packages():
    """Автоматически устанавливает необходимые библиотеки."""
    required = {
        'openpyxl': 'openpyxl',
        'telegram': 'python-telegram-bot'
    }
    
    for module, package in required.items():
        try:
            __import__(module)
            print(f"✅ {module} уже установлен")
        except ImportError:
            print(f"📦 Устанавливаю {package}...")
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", 
                package, "--quiet", "--no-cache-dir"
            ])
            print(f"✅ {package} установлен")

install_packages()

# ============================================
#  ИМПОРТЫ
# ============================================

import json
import logging
import shutil
import time
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.error import RetryAfter, TimedOut, NetworkError, BadRequest
import openpyxl
import asyncio

# ============================================
#  НАСТРОЙКИ
# ============================================

# ⚠️ ЗАМЕНИТЕ НА НОВЫЙ ТОКЕН!
BOT_TOKEN = "8239333710:AAHtiMFQSNoc67WWCcHezW7m-hQ9FjxjfZ8"

# URL Mini App (ваш лендинг)
# Теперь это будет не просто лендинг, а полноценное приложение
MINI_APP_URL = "https://amazing-entremet-0236e3.netlify.app"

# Пути к файлам
SCHEDULE_JSON = 'schedule.json'
FILE_1 = '7d5387db9491af08bdc7e1738e02d263.xlsx'
FILE_2 = 'ae2454c52a6854c7272c28bf32f8d521.xlsx'
ADMIN_FILE = 'admin_id.txt'

# ... (остальные константы DAYS_RU, DAYS_RU_REVERSE, DAYS_SHORT — такие же как раньше)

# ============================================
#  КЛАВИАТУРЫ С MINI APP
# ============================================

def get_main_keyboard():
    """Главная клавиатура с кнопкой открытия Mini App."""
    keyboard = [
        [InlineKeyboardButton(
            "📱 Открыть расписание (Mini App)", 
            web_app=WebAppInfo(url=MINI_APP_URL)
        )],
        [InlineKeyboardButton("📖 Помощь", callback_data="help")]
    ]
    return InlineKeyboardMarkup(keyboard)

def get_groups_keyboard():
    """Клавиатура со списком групп (для тех, кто не хочет открывать Mini App)."""
    groups = get_all_groups()
    
    if not groups:
        keyboard = [[InlineKeyboardButton("⏳ Расписание загружается...", callback_data="noop")]]
        return InlineKeyboardMarkup(keyboard)
    
    keyboard = []
    row = []
    for i, group in enumerate(groups):
        row.append(InlineKeyboardButton(group, callback_data=f"group_{group}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)
    
    # Добавляем кнопку открытия Mini App внизу
    keyboard.append([
        InlineKeyboardButton(
            "📱 Открыть в Mini App", 
            web_app=WebAppInfo(url=MINI_APP_URL)
        )
    ])
    
    return InlineKeyboardMarkup(keyboard)

# ============================================
#  ОБРАБОТЧИКИ КОМАНД
# ============================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start — теперь с Mini App."""
    user_id = update.effective_user.id
    
    # Назначаем администратора
    admin_id = get_admin_id()
    if admin_id is None:
        set_admin_id(user_id)
        await update.message.reply_text(
            f"✅ Вы назначены администратором бота!\n"
            f"Ваш ID: `{user_id}`",
            parse_mode="Markdown"
        )
    
    # Приветствие с кнопкой Mini App
    await update.message.reply_text(
        "👋 *Привет! Я бот расписания экономического факультета.*\n\n"
        "📱 Нажми на кнопку ниже, чтобы открыть *Mini App* с расписанием.\n"
        "Там удобно: выбери группу, неделю и день — и всё расписание перед тобой!\n\n"
        "💡 *Или* просто напиши название своей группы (например, *Э-О-25/1*) — "
        "я покажу расписание на сегодня.",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard()
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Команда /help."""
    await update.message.reply_text(
        "📖 *Как пользоваться ботом*\n\n"
        "📱 *Mini App* — самый удобный способ:\n"
        "   Нажми на кнопку *«Открыть расписание»* и выбери свою группу.\n\n"
        "💬 *Через чат:*\n"
        "   Просто напиши название группы (например, *Э-О-25/1*).\n\n"
        "🔹 *Команды:*\n"
        "   /start — начать\n"
        "   /help — эта справка\n"
        "   /myid — показать мой ID\n"
        "   /update_schedule — обновить расписание (админ)",
        parse_mode="Markdown",
        reply_markup=get_main_keyboard()
    )

# ... (остальные обработчики: show_my_id, update_schedule_command, 
#      button_callback, handle_file_upload, handle_message — такие же как раньше)

# ============================================
#  ЗАПУСК БОТА
# ============================================

def main():
    """Запуск бота."""
    load_schedule()
    
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Регистрируем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("myid", show_my_id))
    application.add_handler(CommandHandler("update_schedule", update_schedule_command))
    application.add_handler(CallbackQueryHandler(button_callback))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_file_upload))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    print("=" * 50)
    print("🤖 БОТ РАСПИСАНИЯ С MINI APP ЗАПУЩЕН")
    print("=" * 50)
    print(f"📱 Mini App URL: {MINI_APP_URL}")
    print("=" * 50)
    
    application.run_polling()

if __name__ == '__main__':
    main()
