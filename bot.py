#!venv/bin/python
# -*- encoding: utf-8 -*-
import config
import logging
import random as rd
import re
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from platform import python_version

import gspread
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils.exceptions import BotBlocked
from aiogram.utils.executor import start_webhook


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return len(str_list) + 1


def check_email(email):
    if email == "admin":
        return True
    return re.fullmatch(r'st\d{6}@student.spbu.ru', email)


def check_lec(text):
    string = text.split('\n')
    for i in string:
        words = i.split("-")
        try:
            if not (len(words) == 2 and int(words[1]) and len(words[1]) > 0):
                return False
        except:
            return False
    return True


def generate_code():
    return rd.randint(1000, 9999)


gc = gspread.service_account(filename='.config/lectop-27ba2c7b13a5.json')
sh = gc.open("LecTOP2023")
worksheet1 = sh.get_worksheet(0)
worksheet2 = sh.get_worksheet(1)

globalrow1 = next_available_row(worksheet1)
globalrow2 = next_available_row(worksheet2)

WEBHOOK_HOST = 'https://pmpu.site'
WEBHOOK_PATH = '/lecTop/'
WEBHOOK_URL = f"{WEBHOOK_HOST}{WEBHOOK_PATH}"

WEBAPP_HOST = '127.0.0.1'
WEBAPP_PORT = 7787

# Объект бота
bot = Bot(token=config.TOKEN)

bak = pd.read_excel("bak.xlsx")

storage = MemoryStorage()
# storage = MongoStorage(host='localhost', port=27017, db_name='aiogram_fsm')
# dp = Dispatcher(bot, storage=storage)
# Диспетчер для бота
dp = Dispatcher(bot, storage=storage)
# Включаем логирование, чтобы не пропустить важные сообщения
logging.basicConfig(level=logging.INFO)




class Form(StatesGroup):
    email = State()
    bak = State()
    answer = State()
    code = State()
    lec = State()
    prac = State()


async def on_startup(dp):
    await bot.set_webhook(WEBHOOK_URL)


async def on_shutdown(dp):
    await bot.delete_webhook()


@dp.message_handler(commands="start")
async def start(message: types.Message):
    await Form.email.set()
    await message.answer(
        "Хэй! Добро пожаловать на LecTOP – голосование за ЛУЧШИХ преподавателей ПМ-ПУ СПбГУ!\nЧтобы перейти, наконец, к выбору, введи свою корпоративную почту в формате: stXXXXXX@student.spbu.ru\nВ любой момент ты можешь начать заново, написав /cancel")


@dp.message_handler(state='*', commands='cancel')
@dp.message_handler(Text(equals='отмена', ignore_case=True), state='*')
async def cancel_handler(message: types.Message, state: FSMContext):
    await state.finish()
    await message.reply('ОК, \cancel так \cancel. Можешь начать сначала написав /start!')


@dp.message_handler(state=Form.prac, commands='back')
@dp.message_handler(state=Form.code, commands='back')
async def back_handler(message: types.Message, state: FSMContext):
    if state == Form.prac:
        await Form.previous()
        return await message.reply(
            "Хорошо, вернемся к предыдущему шагу. Правила те же!")
    else:
        await Form.previous()
        return await message.reply("Хорошо, вернемся к предыдущему шагу. Правила все те же!")


@dp.message_handler(commands="help")
async def help(message: types.Message):
    await message.answer(
        "/start - запуск бота\n/help - основные команды\n/cancel - прервать работу с ботом\n/back - вернуться к прошлому этапу")


@dp.message_handler(lambda message: not check_email(message.text), state=Form.email)
async def process_email_invalid(message: types.Message):
    return await message.reply("Не, немного не то, попробуй ввести именно в таком формате: stXXXXXX@student.spbu.ru")


@dp.message_handler(lambda message: check_email(message.text), state=Form.email)
async def process_email(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['email'] = message.text
    await Form.next()
    if message.text == "admin":
        async with state.proxy() as data:
            data['answer'] = "0000"
            await Form.next()
            data['bak'] = True
            await Form.next()
        return
    if sum(bak['Корпоративный email'] == message.text) == 0:
        await message.answer(
            "Упс! К сожалению, ты – не студент бакалавриата ПМ-ПУ СПбГУ, поэтому твой голос не может быть учтен при подсчете. Но ты все еще можешь испытать бота и проголосовать чисто для себя) Для этого введи сюда код подтверждения, который был выслан на указанную почту.")
        await message.answer("Если ты хочешь вернуться к предыдущему шагу, то напиши /back")
        async with state.proxy() as data:
            data['bak'] = False
        await Form.next()
    else:
        await message.answer(
            "Супер! Ты действительно студент бакалавриата ПМ-ПУ СПбГУ. На указанную почту был выслан код подтверждения, введи его сюда, чтобы окончательно подтвердить свою личность.")
        await message.answer("Если ты хочешь вернуться к предыдущему шагу, то напиши /back")
        async with state.proxy() as data:
            data['bak'] = True
        await Form.next()
    ans = str(generate_code())
    async with state.proxy() as data:
        data['answer'] = ans
    await Form.next()
    recipients = [data['email']]
    sender = 'vikarp21@mail.ru'
    subject = 'LecTOP2023'
    text = 'Код подтверждения авторизации: ' + ans
    html = '<html><head></head><body><p>' + text + '</p></body></html>'
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = 'UNK AMCP SPBU <' + sender + '>'
    msg['To'] = ', '.join(recipients)
    msg['Reply-To'] = sender
    msg['Return-Path'] = sender
    msg['X-Mailer'] = 'Python/' + (python_version())

    part_text = MIMEText(text, 'plain')
    part_html = MIMEText(html, 'html')

    msg.attach(part_text)
    msg.attach(part_html)

    mail = smtplib.SMTP_SSL(config.server)
    mail.login(config.user, config.password)
    mail.sendmail(sender, recipients, msg.as_string())
    mail.quit()


@dp.message_handler(state=Form.code)
async def process_code(message: types.Message, state: FSMContext):
    data = await state.get_data()
    if message.text == data['answer']:
        data['code'] = message.text
        await Form.next()
        await message.reply("Личность подтверждена!")
        await message.answer(
            "Отлично! Теперь можно переходить к голосованию. Для этого вспомни своих любимых преподавателей, их фамилии и обязательно инициалы, иначе твой голос может быть не учтен.\n\nСначала проголосуем за ЛУЧШИХ, по твоему мнению, ЛЕКТОРОВ ПМ-ПУ СПбГУ!")
        await message.answer(
            "Введи свой выбор в одном сообщении строго в представленном формате.\nПример:\n\nИванов А. А. - 20\nАбрамов В. В. - 30\nСтрельцов К. К. - 50")

        return await message.answer("Распределяй баллы с умом. Их сумма не должна превышать 100.")
    return await message.reply("Неправильный код. Повтори попытку.")


@dp.message_handler(lambda message: not check_lec(message.text), state=Form.lec)
async def wrong_lec(message: types.Message, state: FSMContext):
    return await message.answer("Хм… Кажется, ты немного ошибся с форматом, не могу разобрать, попробуй еще раз) ")


@dp.message_handler(lambda message: check_lec(message.text), state=Form.lec)
async def vote_lec(message: types.Message, state: FSMContext):
    global globalrow1
    try:
        lectors = message.text.split('\n')
        point = 0
        for i in range(len(lectors)):
            lectors[i] = lectors[i].split("-")
            point += int(lectors[i][1])
    except:
        return await message.answer("Хм… Кажется, ты немного ошибся с форматом, не могу разобрать, попробуй еще раз) ")
    if point > 100:
        return await message.reply(
            "Подожди, подожди, тут же больше 100 баллов… Так не пойдет, попробуй еще раз, пожалуйста) ")
    async with state.proxy() as data:
        data['lec'] = message.text
    try:
        data = await state.get_data()
        if data['bak']:
            while worksheet1.find(data['email']):
                cell = worksheet1.find(data['email'])
                worksheet1.batch_clear(['A' + str(cell.row) + ":C" + str(cell.row)])
            lectors = message.text.split('\n')
            for i in range(len(lectors)):
                lectors[i] = lectors[i].split("-")
                worksheet1.batch_update([{
                    'range': f'A{globalrow1}:C{globalrow1}',
                    'values': [[data['email'], lectors[i][0], lectors[i][1]]]}])
                globalrow1 += 1
            await message.answer('Твой выбор записан!')
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
        return await message.reply(
            "Аааааа… Прости, что-то пошло не так, видимо я не все предусмотрел, попробуй обратиться к моему создателю https://vk.com/vikarp21, думаю, он сможет решить эту проблему) ")
    await Form.next()
    return await message.answer(
        "Круто! Теперь – за ЛУЧШИХ, по твоему мнению, ПРАКТИКОВ ПМ-ПУ СПбГУ! Правила и формат те же)")


@dp.message_handler(lambda message: not check_lec(message.text), state=Form.prac)
async def wrong_prac(message: types.Message, state: FSMContext):
    return await message.answer("Хм… Кажется, ты немного ошибся с форматом, не могу разобрать, попробуй еще раз) ")


@dp.message_handler(lambda message: check_lec(message.text), state=Form.prac)
async def vote_prac(message: types.Message, state: FSMContext):
    global globalrow2
    try:
        practices = message.text.split('\n')
        point = 0
        for i in range(len(practices)):
            practices[i] = practices[i].split("-")
            point += int(practices[i][1])
    except:
        return await message.answer("Хм… Кажется, ты немного ошибся с форматом, не могу разобрать, попробуй еще раз) ")

    if point > 100:
        return await message.reply(
            "Подожди, подожди, тут же больше 100 баллов… Так не пойдет, попробуй еще раз, пожалуйста) ")
    async with state.proxy() as data:
        data['prac'] = message.text
    data = await state.get_data()
    df = pd.read_csv('result.csv', index_col=0)
    df.loc[len(df.index)] = [data['email'], data['lec'], data['prac']]
    df.to_csv('result.csv')
    try:
        data = await state.get_data()
        if data['bak']:
            while worksheet2.find(data['email']):
                cell = worksheet2.find(data['email'])
                worksheet2.batch_clear(['A' + str(cell.row) + ":C" + str(cell.row)])
            practices = message.text.split('\n')
            for i in range(len(practices)):
                practices[i] = practices[i].split("-")
                worksheet2.batch_update([{
                    'range': f'A{globalrow2}:C{globalrow2}',
                    'values': [[data['email'], practices[i][0], practices[i][1]]]}])
                globalrow2 += 1
            await message.answer('Твой выбор записан!')


    except:
        return await message.reply(
            "Аааааа… Прости, что-то пошло не так, видимо я не все предусмотрел, попробуй обратиться к моему создателю https://vk.com/vikarp21, думаю, он сможет решить эту проблему) ")
    await state.finish()
    return await message.answer(
        'Шик! Спасибо огромное за твой голос и уделенное время! Если вдруг захочешь переголосовать, то просто введи команду /start. Удачного дня)')


@dp.errors_handler(exception=BotBlocked)
async def error_bot_blocked(update: types.Update, exception: BotBlocked):
    # Update: объект события от Telegram. Exception: объект исключения
    # Здесь можно как-то обработать блокировку, например, удалить пользователя из БД
    print(f"Меня заблокировал пользователь!\nСообщение: {update}\nОшибка: {exception}")

    # Такой хэндлер должен всегда возвращать True,
    # если дальнейшая обработка не требуется.
    return True


if __name__ == "__main__":
    # Запуск бота
    # executor.start_polling(dp, skip_updates=True)
    start_webhook(
        dispatcher=dp,
        webhook_path=WEBHOOK_PATH,
        on_startup=on_startup,
        on_shutdown=on_shutdown,
        skip_updates=True,
        host=WEBAPP_HOST,
        port=WEBAPP_PORT,
    )
