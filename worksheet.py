import httplib2
import apiclient.discovery
import ss as ss
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = 'test-proj-for-habr-article-1ab131d98a6b.json'  # имя файла с закрытым ключом

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                  'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

spreadsheet = service.spreadsheets().create(body = {
    'properties': {'title': 'Сие есть название документа', 'locale': 'ru_RU'},
    'sheets': [{'properties': {'sheetType': 'GRID',
                               'sheetId': 0,
                               'title': 'Сие есть название листа',
                               'gridProperties': {'rowCount': 8, 'columnCount': 5}}}]
}).execute()

{'properties': {'autoRecalc': 'ON_CHANGE',
                'defaultFormat': {'backgroundColor': {'blue': 1,
                                                      'green': 1,
                                                      'red': 1},
                                  'padding': {'bottom': 2,
                                              'left': 3,
                                              'right': 3,
                                              'top': 2},
                                  'textFormat': {'bold': False,
                                                 'fontFamily': 'arial,sans,sans-serif',
                                                 'fontSize': 10,
                                                 'foregroundColor': {},
                                                 'italic': False,
                                                 'strikethrough': False,
                                                 'underline': False},
                                  'verticalAlignment': 'BOTTOM',
                                  'wrapStrategy': 'OVERFLOW_CELL'},
                'locale': 'ru_RU',
                'timeZone': 'Etc/GMT',
                'title': 'Сие есть название документа'},
 'sheets': [{'properties': {'gridProperties': {'columnCount': 5,
                                               'rowCount': 8},
                            'index': 0,
                            'sheetId': 0,
                            'sheetType': 'GRID',
                            'title': 'Сие есть название листа'}}],
 'spreadsheetId': '1Sfl7EQ0Yuyo65INidt4LCrHMzFI9wrmc96qHq6EEqHM'}

driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth)
shareRes = driveService.permissions().create(
    fileId = spreadsheet['spreadsheetId'],
    body = {'type': 'user', 'role': 'writer', 'emailAddress': 'timstab1@gmail.com'},  # доступ на чтение кому угодно
    fields = 'id'
).execute()

# ss - экземпляр нашего класса Spreadsheet
ss.prepare_setColumnWidth(0, 317)
ss.prepare_setColumnWidth(1, 200)
ss.prepare_setColumnsWidth(2, 3, 165)
ss.prepare_setColumnWidth(4, 100)
ss.runPrepared()

results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheet['spreadsheetId'], body = {
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Сие есть название листа!B2:C3",
         "majorDimension": "ROWS",     # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
         "values": [["This is B2", "This is C2"], ["This is B3", "This is C3"]]},

        {"range": "Сие есть название листа!D5:E6",
         "majorDimension": "COLUMNS",  # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
    ]
}).execute()

# ss - экземпляр нашего класса Spreadsheet
ss.prepare_setValues("B2:C3", [["This is B2", "This is C2"], ["This is B3", "This is C3"]])
ss.prepare_setValues("D5:E6", [["This is D5", "This is D6"], ["This is E5", "=5+5"]], "COLUMNS")
ss.runPrepared()

