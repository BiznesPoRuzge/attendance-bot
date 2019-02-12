import telebot
from telebot import apihelper
import pendulum
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import string

import constants


def parseint(string):
    return int(''.join([x for x in string if x.isdigit()]))


PROXY = 'socks5://471625721:TjhpFr8Q@phobos.public.opennetwork.cc:1090'
apihelper.proxy = {'https': PROXY}
BOT_TOKEN = constants.TOKEN
bot = telebot.TeleBot(BOT_TOKEN)

# авторизация
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
                                           'SeregaAccred-7d72aeb05dd4.json',
                                           scope
                                           )
sleep(1)  # чтоб не упало из-за слишком частых обращений к апи
gc = gspread.authorize(credentials)
sleep(1)

# открываем таблицу
sheet = gc.open_by_key('12RpRLUd_x0VLhllp7MbXcvpnEtj93yqzo97oVWs_mwI')
sleep(1)

months = [i.title for i in sheet.worksheets()]
month_offset = [2, 3, 2, 3, 3, 0, 3, 2, 3, 2, 3, 3]
cols = ['']

if not pendulum.now().year % 4:
    month_offset[5] += 1

data = {
        'picked_month': None,
        'picked_date': None
        }

print('started')


@bot.message_handler(commands=['start'])
def start(message):
    buttons = []
    markup = telebot.types.ReplyKeyboardMarkup(one_time_keyboard=True)
    for i in months:
        buttons.append(telebot.types.KeyboardButton(i))
    markup.add(*buttons)
    msg = bot.send_message(
                     message.chat.id,
                     'Привет! Я - бот, который может вывести посещаемость '
                     'студентов 2 курса ИВТ в РГПУ им. Герцена. Для начала '
                     'выбери месяц, нажав на одну из кнопок.',
                     reply_markup=markup
                     )
    bot.register_next_step_handler(msg, monthpick)


def monthpick(message):
    data['picked_month'] = months.index(message.text)
    markup = telebot.types.ReplyKeyboardRemove(selective=False)
    if message.text not in months:
        bot.send_message(
                         message.chat.id,
                         'Увы, но такого месяца нет в таблице.'
                         )
    else:
        msg = bot.send_message(
                         message.chat.id,
                         'Введи число.',
                         reply_markup=markup
                         )
        bot.register_next_step_handler(msg, datepick)


def datepick(message):
    data['picked_date'] = parseint(message.text)
    if data['picked_date'] > 0:
        if data['picked_date'] < 28 + month_offset[data['picked_month']]:
            sel_sheet = sheet.get_worksheet(data['picked_month'])
            sel_column = sel_sheet.find(str(data['picked_date'])).col
            sel_column = string.ascii_uppercase[sel_column - 1]
            cell_list = sel_sheet.range(f'{sel_column}1:{sel_column}38')
            cell_vals = [i.value for i in cell_list]
            bot.send_message(
                     message.chat.id,
                     '|'.join(cell_vals)
                     )
        else:
            bot.send_message(
                             message.chat.id,
                             'В выбранном месяце нет такого числа!'
                            )
    else:
        bot.send_message(message.chat.id, 'Введённое число меньше нуля!')


bot.polling()
