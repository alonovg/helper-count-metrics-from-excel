import json

from bot.config import bot
from bot.scripts.filemodify import get_file_data
from bot.scripts.saver import save_and_send_statistics, create_data_dir


@bot.message_handler(commands=['start'])
def start_handler(message):
    if message.text == "/start":
        bot.reply_to(message,
                     f"Привет, {message.from_user.first_name},\nЯ помогаю подсчитать данные {bot.get_me().first_name})")
    elif len(message.text.split()) == 2:
        file = message.text.split()[1]
        file_id, file_type, _, file_extension = get_file_data(message)
        try:
            with open(f"data/{file}.{file_extension}", "rb") as file_data:
                file_data = json.loads(file_data.read())
                file_type = file_data['file_type']
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
