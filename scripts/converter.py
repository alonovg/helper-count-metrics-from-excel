import convertapi


def convert_xls_to_xlsx(file_id, response):
    file_path = f'data/{file_id}.xls'

    with open(file_path, "wb") as xls_file:
        xls_file.write(response.content)

    xlsx_file = f"data/{file_id}.xlsx"

    # Конвертируем в формат XLSX
    convertapi.convert('xlsx', {'File': file_path}, from_format='xls').save_files(xlsx_file)

    return xlsx_file
