import os
import matplotlib
matplotlib.use('Agg')
import telebot
from telebot import types
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import matplotlib.pyplot as plt
import io

load_dotenv()
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
if not TOKEN:
    print("ОШИБКА: TELEGRAM_BOT_TOKEN не найден!")
    exit(1)

EXCEL_FILE = r"C:\Users\HomePC\YandexDisk\Bots\CashFlowMaste4_bot\Data.xlsx"
SHEET_NAME = "Доходы и Расходы"

bot = telebot.TeleBot(TOKEN, parse_mode='HTML')
user_state = {}
last_added_row = {}  # ХРАНИМ ИНДЕКС ПОСЛЕДНЕЙ ДОБАВЛЕННОЙ СТРОКИ

def get_df():
    return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

def save_df(df):
    try:
        df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
    except PermissionError:
        raise PermissionError("Закрой Excel!")

def current_month_name():
    return datetime.now().strftime("%B %Y").replace("November", "ноябрь").replace("October", "октябрь")

def get_months():
    df = get_df()
    months = set()
    for val in df.iloc[:, 1]:
        if isinstance(val, str) and '2025' in val:
            months.add(val)
    return sorted(months, reverse=True)

def get_stats(month_name):
    df = get_df()
    found = False
    income = expense = 0
    for _, row in df.iterrows():
        if str(row.iloc[1]) == month_name:
            found = True
            continue
        if not found: continue
        if pd.isna(row.iloc[0]): break
        if row.iloc[1] == 'Доход': income += row.iloc[2]
        elif row.iloc[1] == 'Расход': expense += row.iloc[2]
    return income, expense

def show_main_menu(chat_id):
    markup = types.InlineKeyboardMarkup(row_width=1)
    markup.add(
        types.InlineKeyboardButton("Добавить", callback_data="main_add"),
        types.InlineKeyboardButton("Отменить добавление", callback_data="main_cancel"),
        types.InlineKeyboardButton("Статистика за месяц", callback_data="main_stats")
    )
    bot.send_message(chat_id, "<b>Выберите действие:</b>", reply_markup=markup)

@bot.message_handler(commands=['start'])
def start(message):
    bot.delete_message(message.chat.id, message.message_id)
    bot.send_message(message.chat.id, "Йо, братан! Готов вести финансы как профи?\n\n/add — добавить\n/logs — логи\n/stats — статистика")
    show_main_menu(message.chat.id)

@bot.message_handler(commands=['add'])
def add_start(message):
    bot.delete_message(message.chat.id, message.message_id)
    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton("Доход", callback_data="inc"),
        types.InlineKeyboardButton("Расход", callback_data="exp")
    )
    bot.send_message(message.chat.id, "Тип транзакции:", reply_markup=markup)
    user_state[message.chat.id] = {"step": "type"}

@bot.message_handler(commands=['stats'])
def stats_current(message):
    bot.delete_message(message.chat.id, message.message_id)
    send_stats(message.chat.id, current_month_name())

@bot.message_handler(commands=['logs'])
def logs_current(message):
    bot.delete_message(message.chat.id, message.message_id)
    send_logs(message.chat.id, current_month_name())

def send_stats(chat_id, month):
    income, expense = get_stats(month)
    fig, ax = plt.subplots(figsize=(8, 6))
    bars = ax.bar(['Доходы', 'Расходы'], [income, expense], color=['#36A2EB', '#FF6384'])
    ax.set_title(f"Статистика: {month}", fontsize=16, fontweight='bold')
    ax.bar_label(bars, fmt='{:,.0f}', fontsize=14, fontweight='bold')
    plt.tight_layout()
    bio = io.BytesIO()
    plt.savefig(bio, format='png')
    bio.seek(0)
    plt.close('all')

    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Другой месяц", callback_data="choose_stats"))

    bot.send_photo(chat_id, bio, caption=f"<b>{month}</b>\nДоходы: {income:,}\nРасходы: {expense:,}", reply_markup=markup)

def send_logs(chat_id, month):
    df = get_df()
    found = False
    lines = [f"<b>Логи: {month}</b>"]
    for _, row in df.iterrows():
        if str(row.iloc[1]) == month:
            found = True
            continue
        if not found: continue
        if pd.isna(row.iloc[0]): break
        line = f"{int(row.iloc[0]):2d} | {row.iloc[1]:6} | {int(row.iloc[2]):8,} | {row.iloc[3]:7} | {row.iloc[4]:10} | {row.iloc[5]}"
        lines.append(line)
    
    text = "\n".join(lines) if len(lines) > 1 else "Нет записей"
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton("Другой месяц", callback_data="choose_logs"))
    
    bot.send_message(chat_id, f"<pre>{text}</pre>", reply_markup=markup)

# === ГЛАВНОЕ МЕНЮ ===
@bot.callback_query_handler(func=lambda call: call.data in ["main_add", "main_stats", "main_cancel"])
def main_menu_handler(call):
    bot.answer_callback_query(call.id)
    if call.data == "main_add":
        add_start(call.message)
    elif call.data == "main_stats":
        stats_current(call.message)
    elif call.data == "main_cancel":
        cancel_last_transaction(call)

# === ОТМЕНА ПОСЛЕДНЕЙ ===
@bot.callback_query_handler(func=lambda call: call.data == "main_cancel")
def cancel_last_transaction(call):
    bot.answer_callback_query(call.id)
    chat_id = call.message.chat.id

    if chat_id not in last_added_row:
        bot.send_message(chat_id, "Нечего отменять — ты ничего не добавлял.")
        show_main_menu(chat_id)
        return

    try:
        df = get_df()
        idx = last_added_row[chat_id]
        if idx >= len(df):
            bot.send_message(chat_id, "Ошибка: запись уже удалена.")
            del last_added_row[chat_id]
            show_main_menu(chat_id)
            return

        removed = df.iloc[idx]
        if removed.iloc[4] != call.from_user.first_name:
            bot.send_message(chat_id, "Это не твоя запись!")
            show_main_menu(chat_id)
            return

        df = df.drop(idx).reset_index(drop=True)
        save_df(df)
        del last_added_row[chat_id]

        sign = "+" if removed.iloc[1] == "Доход" else "-"
        bot.send_message(chat_id, f"<b>УДАЛЕНО!</b>\n{sign}{int(removed.iloc[2]):,} — {removed.iloc[5]}")
        
    except PermissionError:
        bot.send_message(chat_id, "Закрой Excel!")
    except Exception as e:
        bot.send_message(chat_id, f"Ошибка: {e}")
    
    show_main_menu(chat_id)

# === ДОБАВЛЕНИЕ ===
@bot.callback_query_handler(func=lambda call: call.data in ["inc", "exp"])
def set_type(call):
    bot.answer_callback_query(call.id)
    typ = "Доход" if call.data == "inc" else "Расход"
    user_state[call.message.chat.id] = {"step": "amount", "type": typ}
    bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=f"Выбран: <b>{typ}</b>\n\nВведи сумму:")

@bot.message_handler(func=lambda m: user_state.get(m.chat.id, {}).get("step") == "amount")
def get_amount(message):
    try:
        amount = float(message.text.replace(',', '.').replace(' ', ''))
        user_state[message.chat.id].update({"amount": abs(amount), "step": "note"})
        bot.send_message(message.chat.id, "Примечание (или просто отправь):")
    except:
        bot.send_message(message.chat.id, "Введи число!")

@bot.message_handler(func=lambda m: user_state.get(m.chat.id, {}).get("step") == "note")
def save_note(message):
    chat_id = message.chat.id
    state = user_state[chat_id]
    note = message.text.strip() if message.text.strip() else "—"
    
    try:
        df = get_df()
        current_month = current_month_name()
        if not any(str(x) == current_month for x in df.iloc[:, 1]):
            header = pd.DataFrame([['', current_month, '', '', '', '']], columns=df.columns)
            df = pd.concat([df, header], ignore_index=True)
        
        new_row = {
            'Дата': datetime.now().day,
            'Транзакция': state["type"],
            'Сумма': state["amount"],
            'Источник': 'Наличка',
            'Добавил': message.from_user.first_name,
            'Примечание': note
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_df(df)
        
        last_added_row[chat_id] = len(df) - 1  # ЗАПОМИНАЕМ ИНДЕКС!
        
        sign = "+" if state["type"] == "Доход" else "-"
        bot.send_message(chat_id, f"<b>ГОТОВО!</b>\n{sign}{state['amount']:,} — {note}")
        
    except PermissionError:
        bot.send_message(chat_id, "Закрой Excel!")
    except Exception as e:
        bot.send_message(chat_id, f"Ошибка: {e}")
    
    del user_state[chat_id]
    show_main_menu(chat_id)

# === ВЫБОР МЕСЯЦА ===
@bot.callback_query_handler(func=lambda call: call.data in ["choose_stats", "choose_logs"])
def choose_month(call):
    bot.answer_callback_query(call.id)
    months = get_months()
    markup = types.InlineKeyboardMarkup(row_width=2)
    prefix = "s_" if call.data == "choose_stats" else "l_"
    for m in months:
        markup.add(types.InlineKeyboardButton(m, callback_data=f"{prefix}{m}"))
    
    text = "Выбери месяц для статистики:" if call.data == "choose_stats" else "Выбери месяц для логов:"
    
    # РЕДАКТИРУЕМ ТОЛЬКО ТЕКСТОВЫЕ СООБЩЕНИЯ!
    try:
        bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=text, reply_markup=markup)
    except:
        bot.send_message(call.message.chat.id, text, reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith(("s_", "l_")))
def show_month_data(call):
    bot.answer_callback_query(call.id)
    month = call.data[2:]
    if call.data.startswith("s_"):
        send_stats(call.message.chat.id, month)
    else:
        send_logs(call.message.chat.id, month)

# === ЗАПУСК ===
print("БОТ ЗАПУЩЕН! ВСЁ ИДЕАЛЬНО! НИКАКИХ ОШИБОК!")
bot.infinity_polling(none_stop=True)