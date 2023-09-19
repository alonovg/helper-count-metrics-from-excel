from bot.config import bot


def get_file_data(message):
    file_type = message.content_type
    file_id = file_name = file_extension = None
    if file_type == 'document':
        file_id = message.document.file_id
        file_name = message.document.file_name
        file_extension = file_name.split(".")[-1]
        if file_id and file_extension in ('xls', 'xlsx'):  # Проверка расширения файла
            bot.send_message(message.chat.id, "Файл получен, ожидайте.", reply_to_message_id=message.id)
    elif file_type == 'photo':
        bot.send_message(message.chat.id, "Необходим файл")
    elif file_type == 'video':
        bot.send_message(message.chat.id, "Необходим файл")
    return file_id, file_type, file_name, file_extension
