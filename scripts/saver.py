import os
import shutil

import requests

from bot.config import bot, BOT_TOKEN
from bot.scripts.converter import convert_xls_to_xlsx
from bot.scripts.filemodify import get_file_data
from bot.scripts.statistic import process_excel_file, send_statistics


def create_data_dir():
    try:
        os.mkdir('data')
    except FileExistsError:
        pass


def save_and_send_statistics(message):
    admins = []
    with open("admin_list.txt", "r", encoding="utf-8") as file:
        admins.append(file.read().splitlines())
    if str(message.from_user.id) not in admins[0]:
        bot.send_message(message.chat.id, "У вас нет прав доступа к этому чату.")
        return

    file_id, _, _, file_extension = get_file_data(message)

    file_path = f"data/{file_id}"

    # Получаем ссылку на файл и скачиваем его
    xls_file_data = bot.get_file(file_id)
    file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{xls_file_data.file_path}"
    response = requests.get(file_url, timeout=60)

    # Скачиваем файл
    with open(file_path, "wb") as file:
        file.write(response.content)

    if file_extension == 'xls':
        xlsx_file_path = convert_xls_to_xlsx(file_id, response)
        statistics = process_excel_file(xlsx_file_path)
        send_statistics(message, statistics)

        if os.path.isfile(xlsx_file_path):
            os.remove(f"{file_path}.xls")
            with open(xlsx_file_path, "rb") as xlsx_file:
                bot.send_document(message.chat.id, xlsx_file)

        # Удаляем файлы
        os.remove(xlsx_file_path)
    else:
        shutil.copyfile(file_path, f'{file_path}.xlsx')
        statistics = process_excel_file(f"{file_path}.xlsx")
        with open(file_path, "rb") as xlsx_file:
            bot.send_document(message.chat.id, xlsx_file)
        send_statistics(message, statistics)

        # Удаляем файлы
        os.remove(file_path)
        os.remove(f'data/{file_id}.xlsx')
