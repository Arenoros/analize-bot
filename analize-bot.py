import os
import pdfplumber
import re
# Открыть PDF
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import configparser
from prettytable import PrettyTable

blood_components = [
    "Лейкоциты",
    "Эритроциты",
    "Гемоглобин",
    "Тромбоциты",
    "Палочкоядерные",
    "Сегментоядерные",
    "Эозинофилы",
    "Лимфоциты",
    "Моноциты",
    "СОЭ",
    "АКН"
]
ROWS_COUNT = len(blood_components) + 1 # +1 для даты
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

date_pattern = re.compile(r'\d{2}\.\d{2}\.\d{4}', re.UNICODE)
pattern = re.compile(r'\w+', re.UNICODE)

config = configparser.ConfigParser()
# Читаем ini-файл
config.read("config.ini")
GAPI_PKEY       = config['Bot']["gapi_json_key"]
SPREADSHEET_ID  = config['Bot']["g_table_id"]
SHEET_ID        = str(config['Bot']["g_sheet_id"])

def insert_v1(values):
    # Настройка доступа
    scope = ["https://spreadsheets.google.com/feeds"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(GAPI_PKEY, SCOPES)
    client = gspread.authorize(credentials)
    print("Авторизация прошла успешно!")
    # Открытие таблицы
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    print("Таблица открыта!")
    sheet = spreadsheet.sheet1  # Выберите лист (например, первый)
    pos =  len(sheet.row_values(1))
    print(pos)
    new_column_data = [
        values,
    ] # Данные для добавления
    sheet.insert_cols(new_column_data, pos)
    print("Данные успешно добавлены!")

def insert_v2(values):
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(GAPI_PKEY, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        #print("Авторизация прошла успешно!")
        # Отправка запроса
        spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = spreadsheet_metadata.get('sheets', [])
        if not sheets:
            return "Таблица не найдена"
        
        sheet_name=''
        for sheet in sheets:
            cur = str(sheet.get('properties', {}).get('sheetId', ''))
            if cur == SHEET_ID:
                sheet_name = sheet['properties']['title']
                break
        
        if not sheet_name:
            return "Лист не найден"
        
        response = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A1:ZZZ1"
        ).execute()

        if "values" not in response:
            return "Таблица для вставки не найдена"
        
        dates = response["values"][0]
        if values[0] in dates:
            return "Данные за эту дату уже добавлены"
        
        pos = len(response["values"][0])
        base_range = {
                "sheetId": SHEET_ID,  # ID листа (обычно 0 для первого листа)
                "startRowIndex": 0,  # Номер строки (0-индексация)
                "endRowIndex": ROWS_COUNT,  # Конец вставляемых данных
                "startColumnIndex": 1,  # Номер колонки
                "endColumnIndex": 2  # Конец вставляемых данных
        }
        new_column_range = {
            "sheetId": SHEET_ID,  # ID листа (обычно 0 для первого листа)
            "startRowIndex": 0,  # Номер строки (0-индексация)
            "endRowIndex": ROWS_COUNT,  # Конец вставляемых данных
            "startColumnIndex": pos,  # Номер колонки
            "endColumnIndex": pos + 1  # Конец вставляемых данных
        }
        rows = []
        for value in values:
            rows.append({"values": [{"userEnteredValue": {"stringValue": value}}]})

        data = {
            "requests": [
                 {    
                    "insertDimension": {
                        "range": {
                            "sheetId": SHEET_ID,  # Замените на ваш sheetId
                            "dimension": "COLUMNS",  # Указываем, что добавляем столбец
                            "startIndex": pos,  # Индекс, начиная с которого добавить (нумерация с 0)
                            "endIndex": pos + 1   # Индекс, до которого добавить (в данном случае добавляем 1 столбец)
                        },
                        "inheritFromBefore": True  # Унаследовать форматирование от предыдущего столбца (True или False)
                    }
                },
                # Копирование форматирования
                {
                    "copyPaste": {
                        "source": base_range,
                        "destination": new_column_range,
                        "pasteType": "PASTE_NORMAL",
                        "pasteOrientation": "NORMAL",
                    }
                },
                # Вставка данных
                {
                    "updateCells": {
                        "range": {
                            "sheetId": SHEET_ID,  # ID листа (обычно 0 для первого листа)
                            "startRowIndex": 0,  # Номер строки (0-индексация)
                            "endRowIndex": len(values),  # Конец вставляемых данных
                            "startColumnIndex": pos,  # Номер колонки
                            "endColumnIndex": pos + 1  # Конец вставляемых данных
                        },
                        "rows": rows,
                        "fields": "userEnteredValue"
                    }
                }
            ]
        }
        #print(rows)
        # Отправка запроса
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=data
        ).execute()
        return "Данные добавлены в таблицу!"
    except Exception as e:
        return f"Произошла ошибка: {e}"
    


async def load_analyze(pdf_path) -> str:
    results=dict()
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) == 0:
            return "No pages found in the PDF file"
        page = pdf.pages[0]
        date = None
        for line in page.extract_text().splitlines():
            match = date_pattern.fullmatch(line)
            if match:
                date = match.group()
                break
        # Извлечение таблиц с текущей страницы
        tables = page.extract_tables()
        for table in tables:
            for row in table[1:]:
                if len(row) == 4:
                    match = pattern.search(row[0])
                    if match and len(row[1]) > 0:
                        results[match.group()] = (row[1], row[3]) #float(row[1].replace(',', '.'))

    akn = float(results["Лейкоциты"][0].replace(',', '.'))/100*(int(results["Палочкоядерные"][0])+int(results["Сегментоядерные"][0]))
    values = [date]
    message = f'Данные от {date}: \n'

    table = PrettyTable()
    table.align = "l"
    table.field_names = ["Параметр", "Значение", "Норма"]

    # Добавляем строки

    for component in blood_components: 
        if component in results:
            values.append(results[component][0])
            table.add_row([component, results[component][0], results[component][1]])
            #message += f'{component}: {results[component]}\n'
    
    table.add_row(['АКН', str(akn).replace('.', ','), ''])
    message += f'<code>\n{table}\n</code>\n'

    os.remove(pdf_path)
    message += insert_v2(values)
    return message
    

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Привет! Отправьте мне анализ я загружу его.")

# Функция для обработки документов
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = update.message.document
    if file:
        file_id = file.file_id
        file_name = file.file_name
        await update.message.reply_text(f"Обработка файла {file_name} начата...")
        # Скачиваем файл
        new_file = await context.bot.get_file(file_id)
        path = await new_file.download_to_drive(file_name)
        message = await load_analyze(path)
        print(message)
        # Уведомляем пользователя
        await update.message.reply_html(message)

def main():
    TG_TOKEN = config['Bot']["tg_token"]
    application = Application.builder().token(TG_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.PDF, handle_document))
    print("Бот запущен!")
    # Запускаем бота
    application.run_polling()

if __name__ == "__main__":
    main()