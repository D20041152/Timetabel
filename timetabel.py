import telebot
import openpyxl
from telebot import types
wb = openpyxl.reader.excel.load_workbook(filename="timetabel_1.xlsx", data_only="True")
token = "1402814246:AAHgAo7tZQYLWu2B_yJPNO93J-86sSwkkjw"
bot = telebot.TeleBot(token)
global head


@bot.message_handler(content_types= ["text"])
def any_msg(message):
    keyboardmain = types.InlineKeyboardMarkup(row_width=1)
    first_button = types.InlineKeyboardButton(text = "Технологический 📐", callback_data = "т")
    second_button = types.InlineKeyboardButton(text="Естественно-научный 🦠", callback_data="е")
    keyboardmain.add(first_button, second_button)
    bot.send_message(message.chat.id, f"Выберите свой профиль, {message.from_user.first_name}", reply_markup=keyboardmain)

@bot.callback_query_handler(func = lambda call:True)


def callback_inline(call):
    if call.data == "mainmenu":
        keyboardmain = types.InlineKeyboardMarkup(row_width=1)
        first_button = types.InlineKeyboardButton(text="Технологический📐🧮🔍", callback_data="т")
        second_button = types.InlineKeyboardButton(text="Естественно-научный🦠🧬🍀", callback_data="е")
        keyboardmain.add(first_button, second_button)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id= call.message.message_id, text = f"Выберите свой профиль, {call.from_user.first_name}", reply_markup=keyboardmain)


    if call.data == "е" or call.data == "т":
        global var
        var = call.data
        keyboard = types.InlineKeyboardMarkup(row_width=2)
        day1 = types.InlineKeyboardButton(text = "🔸 Понедельник", callback_data= "1")
        day2 = types.InlineKeyboardButton(text="🔸 Вторник", callback_data="2")
        day3 = types.InlineKeyboardButton(text="🔸 Среда", callback_data="3")
        day4 = types.InlineKeyboardButton(text="🔸 Четверг", callback_data="4")
        day5 = types.InlineKeyboardButton(text="🔸 Пятница", callback_data="5")
        day6 = types.InlineKeyboardButton(text="🔸 Суббота", callback_data="6")
        day7 = types.InlineKeyboardButton(text="🔸 Воскресенье", callback_data="7")
        backbutton = types.InlineKeyboardButton(text="⬅Назад", callback_data="mainmenu")
        keyboard.add(day1,day2, day3,day4, day5, day6, day7)
        keyboard.add(backbutton)
        bot.edit_message_text(chat_id=call.message.chat.id, message_id= call.message.message_id, text = "Выберите день недели", reply_markup=keyboard)


    if var == "т" :
        wb.active = 1
    elif var == "е":
        wb.active = 2

    global list_1
    list_1 = wb.active

    global head
    head = list_1['A1'].value + list_1['B1'].value + list_1['C1'].value

    if var == "т":
        if call.data == "1" or call.data == "2" or call.data == "3" or call.data == "4" or call.data == "5" or call.data == "6":
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=f"{head}")
        if call.data == "1":
            for i in range(3, 12):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "2":
            for i in range(13, 22):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "3":
            for i in range(23, 32):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "4":
            for i in range(33, 42):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "5":
            for i in range(43, 52):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "6":
            for i in range(53, 62):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "7":
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text = "ВЫХОДНОЙ")

    elif var == "е":
        if call.data == "1" or call.data == "2" or call.data == "3" or call.data == "4" or call.data == "5" or call.data == "6":
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text=f"{head}")
        if call.data == "1":
            for i in range(3, 12):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "2":
            for i in range(13, 22):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "3":
            for i in range(23, 32):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "4":
            for i in range(33, 42):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "5":
            for i in range(43, 52):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "6":
            for i in range(53, 62):
                bot.send_message(call.from_user.id,
                                 f"{list_1['A' + str(i)].value}  {list_1['B' + str(i)].value}  {list_1['C' + str(i)].value}")

        elif call.data == "7":
            bot.edit_message_text(chat_id=call.message.chat.id, message_id=call.message.message_id, text = "ВЫХОДНОЙ")




if __name__ == "__main__":
    while True:
        try:
            bot.polling(none_stop = True)
        except Exception as e:
            time.sleep(3)
            print(e)
