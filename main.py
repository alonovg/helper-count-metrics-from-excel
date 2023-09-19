import os
import json
import shutil

from openpyxl.reader.excel import load_workbook
import requests
import convertapi
import telebot

BOT_TOKEN = "Your bot token goes here"

convertapi.api_secret = 'Your api secret goes here'

bot = telebot.TeleBot(BOT_TOKEN)


def create_data_dir():
    try:
        os.mkdir('data')
    except FileExistsError:
        pass


def convert_xls_to_xlsx(file_id, response):
    file_path = f'data/{file_id}.xls'

    with open(file_path, "wb") as xls_file:
        xls_file.write(response.content)

    xlsx_file = f"data/{file_id}.xlsx"

    # Конвертируем в формат XLSX
    convertapi.convert('xlsx', {'File': file_path}, from_format='xls').save_files(xlsx_file)

    return xlsx_file


def get_file_data(message):
    file_type = message.content_type
    file_id = None
    file_name = None
    if file_type == 'document':
        file_id = message.document.file_id
        file_name = message.document.file_name
        save_file_id(file_id)
    elif file_type == 'photo':
        bot.send_message(message.chat.id, "Необходим файл")
    elif file_type == 'video':
        bot.send_message(message.chat.id, "Необходим файл")
    return file_id, file_type, file_name


def process_excel_file(file_path):
    file = load_workbook(file_path)
    sheet = file.active

    search = context = region = quiz = brand = 0

    for col in sheet['B']:
        if col.value is not None:
            if "utm_medium=search" in col.value:
                search += 1
            elif "utm_medium=context" in col.value:
                context += 1
            elif "Районы" in col.value:
                region += 1
            elif "utm_medium=cpc" in col.value and "квиз" in col.value:
                quiz += 1
            elif "Брендовые" in col.value:
                brand += 1
    total = search + context + quiz + region + brand
    return search, context, region, quiz, brand, total


def save_and_send_statistics(message):
    admins = open("admin_list.txt", "r").read().splitlines()
    for admin in admins:
        if str(message.from_user.id) == admin:
            file_id, file_type, file_name = get_file_data(message)
            file_extension = file_name.split(".")[-1]
            if file_id and file_extension in ('xls', 'xlsx'):  # Проверка расширения файла
                bot.send_message(message.chat.id, "Файл получен, ожидайте.", reply_to_message_id=message.id)

                file_path = f"data/{file_id}"

                # Получаем ссылку на файл и скачиваем его
                xls_file_data = bot.get_file(file_id)
                file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{xls_file_data.file_path}"
                response = requests.get(file_url)

                # Скачиваем файл
                with open(file_path, "wb") as file:
                    file.write(response.content)

                if file_extension == 'xls':
                    xlsx_file_path = convert_xls_to_xlsx(file_id, response)
                    search, context, region, quiz, brand, total = process_excel_file(xlsx_file_path)
                    send_statistics(message, search, context, region, quiz, brand, total)

                    if os.path.isfile(xlsx_file_path):
                        os.remove(f"{file_path}.xls")
                        bot.send_document(message.chat.id, open(xlsx_file_path, "rb"))

                    # Удаляем файлы
                    os.remove(xlsx_file_path)
                else:
                    shutil.copyfile(file_path, f'{file_path}.xlsx')
                    search, context, region, quiz, brand, total = process_excel_file(f"{file_path}.xlsx")
                    bot.send_document(message.chat.id, open(file_path, "rb"))
                    send_statistics(message, search, context, region, quiz, brand, total)

                    # Удаляем файлы
                    os.remove(file_path)
                    os.remove(f'data/{file_id}.xlsx')


def send_statistics(message, search, context, region, quiz, brand, total):
    bot.send_message(message.chat.id, f"""**
    Автостратегия поиск: {search} шт.
    Автостратегия рся: {context} шт.
    рся квиз: {quiz} шт.
    Районы рся: {region} шт.
    Брендовые: {brand} шт.
    Всего {total}
    **""", reply_to_message_id=message.id)


def save_file_id(file_id):
    with open("saved_file_ids.txt", "a") as file:
        file.write(file_id + "\n")


@bot.message_handler(commands=['start'])
def start_handler(message):
    if message.text == "/start":
        bot.reply_to(message,
                     f"Привет, {message.from_user.first_name},\nЯ помогаю подсчитать данные {bot.get_me().first_name})")
    elif len(message.text.split()) == 2:
        file = message.text.split()[1]
        file_id, file_type, file_name = get_file_data(message)
        file_extension = file_name.split(".")[-1]
        try:
            with open(f"data/{file}.{file_extension}", "rb") as file_data:
                file_data = json.loads(file_data.read())
                file_type = file_data['file_type']
                file_id = file_data['file_id']
                if file_type == 'document':
                    bot.send_document(message.chat.id, file_id, reply_to_message_id=message.id)
                elif file_type == 'photo':
                    bot.send_message(message.chat.id, "Необходим файл")
                elif file_type == 'video':
                    bot.send_message(message.chat.id, "Необходим файл")
        except FileNotFoundError:
            bot.reply_to(message, "File not found")


@bot.message_handler(content_types=["photo", "video", "document"])
def add_file(message):
    create_data_dir()
    save_and_send_statistics(message)


if __name__ == '__main__':
    print("Started")
    bot.infinity_polling()
