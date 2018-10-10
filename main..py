# -*- charset=utf-8 -*-

"""
    Создано на основе статьи https://habr.com/post/305378/
"""
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

# Закомментировать для создания нового документа
SHEET_ID = '15IOHG1jzhgchNBwjj1kjb_c4qhwIChdd4VXXz5Xpi0Q'

# Надо юзать свой конфиг. В консоли developers.google.com/console должны быть включены API для Drive и Sheets.
CREDENTIALS_FILE = 'SheetsTest-9b6c88f7dfa1.json'

def main():
    """
        Google Drive необходим для сохранения и доступа к созданному документу.
        Второй параметр - скопы для текущих credentials.
    """
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive'])

    httpAuth = credentials.authorize(httplib2.Http())

    # Сервисы билдятся для каждого API (Drive, Maps, Sheets, etc.)
    service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)


    # Либо тащим существующий док, либо мутим новый
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId = SHEET_ID).execute()
    except NameError:
        spreadsheet = service.spreadsheets().create(body = {
            'properties' : { 
                'title': u'Мое название документа', 
                'locale' : 'ru_RU'
            },
            'sheets':[
                {
                    'properties': {
                        'sheetType': 'GRID',
                        'sheetId':0,
                        'title': u'Мое название листа',
                        'gridProperties': {
                            'rowCount': 8,
                            'columnCount':5
                        }
                    }
                },
                {
                    'properties': {
                        'sheetType': 'GRID',
                        'sheetId':1,
                        'title': u'Второе название листа',
                        'gridProperties': {
                            'rowCount': 8,
                            'columnCount':5
                        }
                    }
                }
            ]
        }).execute()
        # Вывод ссылки на созданный документ
        print("https://docs.google.com/spreadsheets/d/{0}/edit".format(spreadsheet['spreadsheetId']))

    driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth)

    """
        Добавление прав на чтение для всех.
        Для определенного пользователя - {'type': 'user', 'role': 'writer', 'emailAddress': 'user@example.com'}
    """
    shareRes = driveService.permissions().create(
        fileId = spreadsheet['spreadsheetId'],
        body = {
            'type':'anyone',
            'role':'reader'
        },
        fields = 'id'
    ).execute()

    """
        Добавление данных в таблицу.
        Поле range должно совпадать с именами созданных листов. Иначе все зафейлится т.к. не найдется листов.
    """
    results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheet['spreadsheetId'], body = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": u'Мое название листа',
            "majorDimension": "ROWS",     # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
            "values": [["This is B2", "This is C2"], ["This is B3", "This is C3"]]},

            {"range": u'Второе название листа',
            "majorDimension": "COLUMNS",  # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
            "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
        ]
    }).execute()

if __name__ == '__main__':
    main()
    exit(0)