import telebot
from telebot import types
import ast
import openpyxl
import os.path
from datetime import datetime
from Config import TELEGRAM_BOT_API

5FBot = telebot.TeleBot(TELEGRAM_BOT_API)
GloveQuestions = [
    ('Glove1.jpeg', 'Напишіть приблизну кількість продажу заливних рукавиць'),
    ('Glove2.jpeg', 'Напишіть приблизну кількість продажу латексних рукавиць (як у прибиральниць)'),
    ('Glove3.jpeg', 'Напишіть приблизну кількість продажу нитрилових рукавиць (як у лікарів)'),
    ('Glove4.jpeg', 'Напишіть приблизну кількість продажу бавовняно джинсових рукавиць (під великий палець)'),
    ('Glove5.jpeg', 'Напишіть приблизну кількість продажу брезентових (під великий палець)'),
    ('Glove6.jpeg', 'Напишіть приблизну кількість продажу крагі зварювальників зі спілки')
]

FFResult=open('5fResult.txt', 'r').read()
if len(FFResult)==0 or FFResult[0]!='{':
    FiveFingersInfo = {}
else:
    FiveFingersInfo = ast.literal_eval(FFResult)

@5FBot.message_handler(commands=['start'])
def StartChat(Message):
    if Message.chat.username not in ['dmitrokds', 'ksb2006']:
        if Message.chat.username not in FiveFingersInfo:
            FiveFingersInfo[Message.chat.username] = []
        FiveFingersInfo[Message.chat.username].append({'City':'', 'Characteristic':'', 'Glove1':'', 'Glove2':'', 'Glove3':'', 'Glove4':'', 'Glove5':'', 'Glove6':'', 'AdditionalInfo':'', 'DateStart':str(datetime.now().strftime("%Y %m %d %H %M %S")), 'DateEnd':'', 'FirstName':getattr(Message.chat, 'first_name', ''), 'SecondName':getattr(Message.chat, 'last_name', '')})
        MfestBot.send_message(Message.chat.id, f'Привіт {Message.chat.username}')
        MfestBot.send_message(Message.chat.id, f'{Message.chat.username} напиши місто вашої торгової точки')
        MfestBot.register_next_step_handler(Message, CityCallback)
    else:
        ButtonsMarkup = types.InlineKeyboardMarkup()
        ButtonsMarkup.row(types.InlineKeyboardButton('Почати тест', callback_data='StartTest'))
        ButtonsMarkup.row(types.InlineKeyboardButton('Отримати статистику', callback_data='GetStatistic'))
        MfestBot.send_message(Message.chat.id, f'Привіт {Message.chat.username}', reply_markup=ButtonsMarkup)


def CityCallback(Message):
    FiveFingersInfo[Message.chat.username][-1]['City'] = Message.text
    ButtonsMarkup = types.InlineKeyboardMarkup()
    ButtonsMarkup.row(types.InlineKeyboardButton('Кількість ниток', callback_data='Кількість нитокSetAnswer'))
    ButtonsMarkup.row(types.InlineKeyboardButton('Bara', callback_data='BaraSetAnswer'))
    ButtonsMarkup.row(types.InlineKeyboardButton("Knac B'93KM 7, 10, 13", callback_data="Knac B'93KM 7, 10, 13SetAnswer"))
    ButtonsMarkup.row(types.InlineKeyboardButton('Все вище перелічене', callback_data='Все вище переліченеSetAnswer'))
    ButtonsMarkup.row(types.InlineKeyboardButton('Ваш варіант', callback_data='SetOwnAnswer'))
    MfestBot.send_message(Message.chat.id, f'На які характеристики вʼязаних рукавиць Клієнт звертає найбільшу увагу при замовленні?', reply_markup=ButtonsMarkup)

def SetOwnAnswerCallback(Message):
    FiveFingersInfo[Message.chat.username][-1]['Characteristic'] = Message.text
    MfestBot.send_message(Message.chat.id, f'Дякую за вашу відповідь.')
    GlovesQuestion(Message, GloveQuestions[0][0], GloveQuestions[0][1])

def GlovesQuestion(Message, GlovePhoto, GloveQuestion):
    MfestBot.send_photo(Message.chat.id, photo=open(GlovePhoto, 'rb'))
    MfestBot.send_message(Message.chat.id, GloveQuestion)
    MfestBot.register_next_step_handler(Message, lambda Message:GlovesAnswer(Message, GlovePhoto, GloveQuestion))

def GlovesAnswer(Message, GlovePhoto, GloveQuestion):
    FiveFingersInfo[Message.chat.username][-1][GlovePhoto.replace('.jpeg', '')] = Message.text
    GloveIndex=GloveQuestions.index((GlovePhoto, GloveQuestion))+1
    if GloveIndex<len(GloveQuestions):
        GlovesQuestion(Message, GloveQuestions[GloveIndex][0], GloveQuestions[GloveIndex][1])
    else:
        FiveFingersInfo[Message.chat.username][-1]['DateEnd'] = str(datetime.now().strftime("%Y %m %d %H %M %S"))
        open('5fResult.txt', 'w').write(str(FiveFingersInfo))
        FiveFingersInfoFile = openpyxl.Workbook()
        FiveFingersInfoColumn = FiveFingersInfoFile.active

        FiveFingersInfoColumn['A1'] = 'Name'
        FiveFingersInfoColumn['B1'] = 'City'
        FiveFingersInfoColumn['C1'] = 'Characteristic'
        FiveFingersInfoColumn['D1'] = 'Glove1'
        FiveFingersInfoColumn['E1'] = 'Glove2'
        FiveFingersInfoColumn['F1'] = 'Glove3'
        FiveFingersInfoColumn['G1'] = 'Glove4'
        FiveFingersInfoColumn['H1'] = 'Glove5'
        FiveFingersInfoColumn['I1'] = 'Glove6'
        FiveFingersInfoColumn['J1'] = 'AdditionalInfo'
        FiveFingersInfoColumn['K1'] = 'DateStart'
        FiveFingersInfoColumn['L1'] = 'DateEnd'
        FiveFingersInfoColumn['M1'] = 'FirstName'
        FiveFingersInfoColumn['N1'] = 'SecondName'

        FiveFingerColumn=2
        for FiveFingerName, FiveFingerInfos in FiveFingersInfo.items():
            for FiveFingerInfo in FiveFingerInfos:
                if FiveFingerInfo != {}:
                    FiveFingersInfoColumn['A'+str(FiveFingerColumn)] = FiveFingerName
                    FiveFingersInfoColumn['B'+str(FiveFingerColumn)] = FiveFingerInfo['City']
                    FiveFingersInfoColumn['C'+str(FiveFingerColumn)] = FiveFingerInfo['Characteristic']
                    FiveFingersInfoColumn['D'+str(FiveFingerColumn)] = FiveFingerInfo['Glove1']
                    FiveFingersInfoColumn['E'+str(FiveFingerColumn)] = FiveFingerInfo['Glove2']
                    FiveFingersInfoColumn['F'+str(FiveFingerColumn)] = FiveFingerInfo['Glove3']
                    FiveFingersInfoColumn['G'+str(FiveFingerColumn)] = FiveFingerInfo['Glove4']
                    FiveFingersInfoColumn['H'+str(FiveFingerColumn)] = FiveFingerInfo['Glove5']
                    FiveFingersInfoColumn['I'+str(FiveFingerColumn)] = FiveFingerInfo['Glove6']
                    FiveFingersInfoColumn['J'+str(FiveFingerColumn)] = FiveFingerInfo['AdditionalInfo']
                    FiveFingersInfoColumn['K'+str(FiveFingerColumn)] = FiveFingerInfo['DateStart']
                    FiveFingersInfoColumn['L'+str(FiveFingerColumn)] = FiveFingerInfo['DateEnd']
                    FiveFingersInfoColumn['M'+str(FiveFingerColumn)] = FiveFingerInfo['FirstName']
                    FiveFingersInfoColumn['N'+str(FiveFingerColumn)] = FiveFingerInfo['SecondName']
                    FiveFingerColumn+=1


        FiveFingersInfoFile.save('FFResult.xlsx')

        ButtonsMarkup = types.InlineKeyboardMarkup()
        ButtonsMarkup.row(types.InlineKeyboardButton('✅ Так', callback_data='QuestionYes'))
        ButtonsMarkup.row(types.InlineKeyboardButton('❌ Ні. Почати з початку', callback_data='StartFromBegining'))
        MfestBot.send_message(Message.chat.id, 'Можливо у вас є якась ще інформація, а ми не задали питання', reply_markup=ButtonsMarkup)

def SetQuestion(Message):
    FiveFingersInfo[Message.chat.username][-1]['AdditionalInfo'] = Message.text
    FiveFingersInfo[Message.chat.username][-1]['DateEnd'] = str(datetime.now().strftime("%Y %m %d %H %M %S"))
    ButtonsMarkup = types.InlineKeyboardMarkup()
    ButtonsMarkup.row(types.InlineKeyboardButton('Почати з початку', callback_data='StartFromBegining'))
    MfestBot.send_message(Message.chat.id, f'Дякую за інформацію', reply_markup=ButtonsMarkup)
    open('5fResult.txt', 'w').write(str(FiveFingersInfo))
    FiveFingersInfoFile = openpyxl.Workbook()
    FiveFingersInfoColumn = FiveFingersInfoFile.active

    FiveFingersInfoColumn['A1'] = 'Name'
    FiveFingersInfoColumn['B1'] = 'City'
    FiveFingersInfoColumn['C1'] = 'Characteristic'
    FiveFingersInfoColumn['D1'] = 'Glove1'
    FiveFingersInfoColumn['E1'] = 'Glove2'
    FiveFingersInfoColumn['F1'] = 'Glove3'
    FiveFingersInfoColumn['G1'] = 'Glove4'
    FiveFingersInfoColumn['H1'] = 'Glove5'
    FiveFingersInfoColumn['I1'] = 'Glove6'
    FiveFingersInfoColumn['J1'] = 'AdditionalInfo'
    FiveFingersInfoColumn['K1'] = 'DateStart'
    FiveFingersInfoColumn['L1'] = 'DateEnd'
    FiveFingersInfoColumn['M1'] = 'FirstName'
    FiveFingersInfoColumn['N1'] = 'SecondName'

    FiveFingerColumn=2
    for FiveFingerName, FiveFingerInfos in FiveFingersInfo.items():
        for FiveFingerInfo in FiveFingerInfos:
            if FiveFingerInfo != {}:
                FiveFingersInfoColumn['A'+str(FiveFingerColumn)] = FiveFingerName
                FiveFingersInfoColumn['B'+str(FiveFingerColumn)] = FiveFingerInfo['City']
                FiveFingersInfoColumn['C'+str(FiveFingerColumn)] = FiveFingerInfo['Characteristic']
                FiveFingersInfoColumn['D'+str(FiveFingerColumn)] = FiveFingerInfo['Glove1']
                FiveFingersInfoColumn['E'+str(FiveFingerColumn)] = FiveFingerInfo['Glove2']
                FiveFingersInfoColumn['F'+str(FiveFingerColumn)] = FiveFingerInfo['Glove3']
                FiveFingersInfoColumn['G'+str(FiveFingerColumn)] = FiveFingerInfo['Glove4']
                FiveFingersInfoColumn['H'+str(FiveFingerColumn)] = FiveFingerInfo['Glove5']
                FiveFingersInfoColumn['I'+str(FiveFingerColumn)] = FiveFingerInfo['Glove6']
                FiveFingersInfoColumn['J'+str(FiveFingerColumn)] = FiveFingerInfo['AdditionalInfo']
                FiveFingersInfoColumn['K'+str(FiveFingerColumn)] = FiveFingerInfo['DateStart']
                FiveFingersInfoColumn['L'+str(FiveFingerColumn)] = FiveFingerInfo['DateEnd']
                FiveFingersInfoColumn['M'+str(FiveFingerColumn)] = FiveFingerInfo['FirstName']
                FiveFingersInfoColumn['N'+str(FiveFingerColumn)] = FiveFingerInfo['SecondName']
                FiveFingerColumn+=1


    FiveFingersInfoFile.save('FFResult.xlsx')


@5FBot.callback_query_handler(func=lambda callback: True)
def SetAnswerCallback(Callback):
    if 'SetAnswer' in Callback.data:
        MfestBot.send_message(Callback.message.chat.id, f"Ви вибрали {Callback.data.replace('SetAnswer', '')}. Дякую за вашу відповідь.")
        MfestBot.edit_message_reply_markup(Callback.message.chat.id, Callback.message.message_id, reply_markup=None)
        FiveFingersInfo[Callback.message.chat.username][-1]['Characteristic'] = Callback.data.replace('SetAnswer', '')
        GlovesQuestion(Callback.message, GloveQuestions[0][0], GloveQuestions[0][1])
    elif Callback.data == 'SetOwnAnswer':
        MfestBot.edit_message_reply_markup(Callback.message.chat.id, Callback.message.message_id, reply_markup=None)
        MfestBot.send_message(Callback.message.chat.id, f'Напишіть, будь ласка, ваш варіант')
        MfestBot.register_next_step_handler(Callback.message, SetOwnAnswerCallback)
    elif Callback.data == 'QuestionYes':
        MfestBot.send_message(Callback.message.chat.id, f'Напишіть, будь ласка, ваше повідомлення')
        MfestBot.register_next_step_handler(Callback.message, SetQuestion)
    elif Callback.data == 'StartFromBegining':
        RestartChat(Callback.message)
    elif Callback.data == 'StartTest':
        if Callback.message.chat.username not in FiveFingersInfo:
            FiveFingersInfo[Callback.message.chat.username] = []
        Message=Callback.message
        FiveFingersInfo[Message.chat.username].append({'City':'', 'Characteristic':'', 'Glove1':'', 'Glove2':'', 'Glove3':'', 'Glove4':'', 'Glove5':'', 'Glove6':'', 'AdditionalInfo':'', 'DateStart':str(datetime.now().strftime("%Y %m %d %H %M %S")), 'DateEnd':'', 'FirstName':getattr(Message.chat, 'first_name', ''), 'SecondName':getattr(Message.chat, 'last_name', '')})
        MfestBot.edit_message_reply_markup(Callback.message.chat.id, Callback.message.message_id, reply_markup=None)
        MfestBot.send_message(Callback.message.chat.id, f'{Callback.message.chat.username} напиши місто вашої торгової точки')
        MfestBot.register_next_step_handler(Callback.message, CityCallback)
    elif Callback.data == 'GetStatistic':
        MfestBot.edit_message_reply_markup(Callback.message.chat.id, Callback.message.message_id, reply_markup=None)
        ButtonsMarkup = types.InlineKeyboardMarkup()
        ButtonsMarkup.row(types.InlineKeyboardButton('Почати з початку', callback_data='StartFromBegining'))
        if os.path.isfile('FFResult.xlsx'):
            MfestBot.send_document(Callback.message.chat.id, open('FFResult.xlsx', 'rb'), reply_markup=ButtonsMarkup)
        else:
            MfestBot.send_message(Callback.message.chat.id, 'Файл не існує', reply_markup=ButtonsMarkup)



def RestartChat(Message):
    if Message.chat.username not in ['dmitrokds', 'ksb2006']:
        if Message.chat.username not in FiveFingersInfo:
            FiveFingersInfo[Message.chat.username] = []
        FiveFingersInfo[Message.chat.username].append({'City':'', 'Characteristic':'', 'Glove1':'', 'Glove2':'', 'Glove3':'', 'Glove4':'', 'Glove5':'', 'Glove6':'', 'AdditionalInfo':'', 'DateStart':str(datetime.now().strftime("%Y %m %d %H %M %S")), 'DateEnd':'', 'FirstName':getattr(Message.chat, 'first_name', ''), 'SecondName':getattr(Message.chat, 'last_name', '')})
        MfestBot.send_message(Message.chat.id, f'Привіт {Message.chat.username}')
        MfestBot.send_message(Message.chat.id, f'{Message.chat.username} напиши місто вашої торгової точки')
        MfestBot.register_next_step_handler(Message, CityCallback)
    else:
        ButtonsMarkup = types.InlineKeyboardMarkup()
        ButtonsMarkup.row(types.InlineKeyboardButton('Почати тест', callback_data='StartTest'))
        ButtonsMarkup.row(types.InlineKeyboardButton('Отримати статистику', callback_data='GetStatistic'))
        MfestBot.send_message(Message.chat.id, f'Привіт {Message.chat.username}', reply_markup=ButtonsMarkup)


5FBot.infinity_polling()
