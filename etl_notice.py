#! pip install pyTelegramBotAPI
import pandas as pd
import numpy as np
import datetime
#import smtplib
#from email.mime.text import MIMEText
#from email.mime.multipart import MIMEMultipart
import telebot
import dataframe_image as dfi
import os

# accesses_info
from accesses_storage import bot_id, channel_name, way_to_off_time_table

def send_to_off_time_yes(df, description):
    '''
    Функция отправляет изображение датафрейма и описание в канал переработок
    '''
    # создание изображения датафрейма
    g = dfi.export(df, 'dataframe.png')
    # инициализация телебота
    bot = telebot.TeleBot(bot_id)
    # фото для отправки в бинарном виде
    doc = open('dataframe.png', 'rb')
    # отправка фото в канал
    bot.send_photo(channel_name, doc, caption=description, parse_mode="html")
    # закрываем и удаляем файл
    doc.close()
    os.remove('dataframe.png')

def send_to_off_time_no(m_y):
    '''
    Функция отправляет сообщеине о  том что переработок нет
    '''
    # инициализация телебота
    bot = telebot.TeleBot(bot_id)
    # отправка сообщения в канал
    bot.send_message(channel_name, text = f"В этом месяце <b>({m_y})</b> нет переработок", parse_mode="html")

# чтение файла
df_work = pd.read_excel(way_to_off_time_table)
# создание столбца с месяцем и годом
df_work['m-y'] = df_work['Дата'].apply(lambda x: str(x.month) + '-' + str(x.year))

# создание переменных с текщим месяцм и текущим годом
m_y_now = str(datetime.date.today().month) + '-' + str(datetime.date.today().year)

# проверка на наличие переработок
swith_off_time = False
if m_y_now in df_work['m-y'].to_list():
    swith_off_time = True

# создание таблицы переработок за текущий месяц
if swith_off_time:
    df_off_time_cur = df_work[df_work['m-y'] == m_y_now]
    df_off_time_cur.fillna('-',inplace=True)
    df_for_send = df_off_time_cur.copy()

# создание суммарного количества часов
if swith_off_time:
    timedelta_time_off = datetime.timedelta()
    for i in df_off_time_cur.index:
        hours_time_off = df_off_time_cur['Время переработки'][i].hour
        minutes_time_off = df_off_time_cur['Время переработки'][i].minute
        timedelta_time_off += datetime.timedelta(hours = hours_time_off, minutes = minutes_time_off)

    minutes_delta_time_off = timedelta_time_off.seconds / 60
    hours_sum_time_off = int(minutes_delta_time_off // 60)
    minutes_sum_time_off = int(minutes_delta_time_off % 60)
    time_for_message = f'Общая сумма переработок в этом месяце (<b>{m_y_now}</b>): \n<b>{hours_sum_time_off} часа {minutes_sum_time_off} минут</b>'

# отправка сообщения
if swith_off_time:
    send_to_off_time_yes(df_off_time_cur, time_for_message)
else:
    send_to_off_time_no(m_y_now)