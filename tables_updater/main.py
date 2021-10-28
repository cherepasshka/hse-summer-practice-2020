import openpyxl
import httplib2
import googleapiclient.discovery
import os
from oauth2client.service_account import ServiceAccountCredentials


def connect_to_document(credentials_file):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_file,
        ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    return googleapiclient.discovery.build('sheets', 'v4', http=httpAuth)


def read(sheet_name, from_cell, to_cell):
    range = '{name}!{from_cell}:{to_cell}'.format(name=sheet_name, from_cell=from_cell, to_cell=to_cell)
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range,
        majorDimension='ROWS'
    ).execute()
    return values


def write(sheet_name, from_cell, data_to_write, dimension):
    last_row = int(from_cell[1:]) + len(data_to_write) - 1
    end_cell = '{letter}{row}'.format(letter=chr(ord(from_cell[0]) + len(data_to_write[0]) - 1), row=last_row)
    range = '{name}!{from_cell}:{end_cell}'.format(name=sheet_name, from_cell=from_cell, end_cell=end_cell)
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": range,
                 "majorDimension": dimension,
                 "values": data_to_write},
            ]
        }
    ).execute()
    return values


def READ_FROM_SHEET(sheet_name):
    data = read(sheet_name, "A3", "E400")['values']
    return [x[:7] + [''] * (7 - len(x)) for x in data if len(x) >= 2]


def READ_FROM_EXCEL(file_name):
    book = openpyxl.open(file_name, read_only=True)
    sheet = book[book.sheetnames[0]]
    data = []
    for row in range(2, min(1000, sheet.max_row)):
        if len(sheet[row]) < 5 or sheet[row][1].value is None:
            continue
        full_name = sheet[row][1].value
        region = str(sheet[row][2].value)
        school = str(sheet[row][3].value)
        grade = str(sheet[row][4].value)
        if len(full_name) == 0 or full_name == ' ':
            continue
        data.append([full_name, region, school, grade])
    book.close()
    return data


def names_equal(name1, name2):
    n1 = name1.split()
    n2 = name2.split()
    if len(n2) < 2 or len(n1) < 2:
        return False
    res = (n1[0] == n2[0] and n1[1] == n2[1])
    if len(n1) == 3 and len(n2) == 3:
        return res and n1[2] == n2[2]
    return res


files_at_folder = files = os.listdir()
if 'input.txt' not in files_at_folder:
    print('Ошибка: в текущей директории нет input.txt, создайте его по шаблону в before-usage/input-data-template.txt')
    exit(1)

from_sheet = "Математика"
with open('input.txt', 'r', encoding='utf-8') as f:
    spreadsheet_id, row_range, columns, update_sheet, excel_file, CREDENTIALS_FILE = f.read().split('\n')[:6]
    row_range = map(int, row_range.split(','))

service = connect_to_document(CREDENTIALS_FILE)
print('Сервисный аккаунт подключился к Google документу, начинается чтение из Excel таблицы')
try:
    students = READ_FROM_EXCEL(excel_file)
except:
    print('Данные из таблицы', excel_file, 'прочитать не удалось.')
    print('Проверьте, пожалуйста, что формат столбцов соответствует тому, что указан в инструкции.')
    exit(0)
print('Данные из таблицы', excel_file, "прочитаны, начинается поиск и обновление google таблицы....")
start, end = row_range
inds = [ord(x) - ord('A') for x in columns]
values = read(update_sheet, "A" + str(start), "Z" + str(end + 1))['values']

values = [x[:26] + [''] * (26 - len(x)) for x in values]
count_of_updated = 0
for i in range(len(values)):
    full_name, region, school, grade = [values[i][j] for j in inds]
    if school != '':
        continue
    for s in students:
        if not names_equal(full_name, s[0]) or grade != s[3]:
            continue
        if s[2] != '' and school == '': #school is non empty
            school = s[2]
            count_of_updated += 1
            print('Найдена школа для', full_name)
            break
    if school == '':
        continue
    updated = [full_name, region, school, grade]
    for j in range(len(inds)):
        values[i][inds[j]] = updated[j]
if count_of_updated > 0:
    write(update_sheet, "A" + str(start), values, "ROWS")
print("Количество школьников, для которых найдена информация о школах:", count_of_updated)