from openpyxl.reader.excel import load_workbook

from bot.config import bot


def send_statistics(message, statistics: dict):
    bot.send_message(message.chat.id, f"""**
    Автостратегия поиск: {statistics["search"]} шт.
    Автостратегия рся: {statistics["context"]} шт.
    рся квиз: {statistics["quiz"]} шт.
    Районы рся: {statistics["region"]} шт.
    Брендовые: {statistics["brand"]} шт.
    Всего {statistics["total"]}
    **""", reply_to_message_id=message.id)


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
    stats_dict = {
        "search": search,
        "context": context,
        "region": region,
        "quiz": quiz,
        "brand": brand,
        "total": search + context + quiz + region + brand
    }
    return stats_dict
