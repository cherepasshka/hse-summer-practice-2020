import httplib2
import os
import googleapiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


# Авторизуемся и получаем service — экземпляр доступа к API
def connect(credentials_file):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_file,
        ['https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    return googleapiclient.discovery.build('sheets', 'v4', http=httpAuth)


def have_sheet(name, all_sheets):
    for sheet in all_sheets:
        if sheet['properties']['title'] == name:
            return True
    return False


def add_sheet(sheet_name, row_count, column_count):
    results = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name,
                            "gridProperties": {
                                "rowCount": row_count,
                                "columnCount": column_count
                            }
                        }
                    }
                }
            ]
        }
    ).execute()
    return results


def read(sheet_name, from_cell, to_cell):
    range = '{name}!{from_cell}:{to_cell}'.format(name=sheet_name, from_cell=from_cell, to_cell=to_cell)
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range,
        majorDimension='ROWS'
    ).execute()
    return values


def clear_sheet(sheet_name, start_cell, row_count, column_count):
    last_row = int(start_cell[1:]) + row_count - 1
    end_cell = '{letter}{row}'.format(letter=chr(ord(start_cell[0]) + column_count - 1), row=last_row)
    range = '{name}!{from_cell}:{end_cell}'.format(name=sheet_name, from_cell=start_cell, end_cell=end_cell)
    request = service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range).execute()
    return request


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


def get_counter(data):
    count = {}
    for x in data:
        if x not in count.keys():
            count[x] = 0
        count[x] += 1
    return count


def get_statistics(data):
    count = get_counter(data)
    data = list(set(data))
    data.sort(key=lambda x: (-count[x], x))
    return [[x, count[x]] for x in data]


def select_region(data, regarded_regions, inds):
    '''
    :param data: [place, full name, city or region, grade, status, link in social nets]
    '''
    students = []
    for value in data:
        if value[inds[0]] == '':
            continue
        student = [value[j] for j in inds]
        for region in regarded_regions:
            if region in student[1] or student[1] in region:
                students.append(student)
                break
    count = get_counter([x[2] for x in students])
    students.sort(key=lambda x: (x[1], -count[x[2]], x[2]))
    return students


def prepare_sheet(sheet_name, sheet_list):
    if not have_sheet(sheet_name, sheet_list):
        add_sheet(sheet_name, 1000, 26)
    clear_sheet(sheet_name, 'A1', row_count=1000, column_count=26)


files_at_folder = files = os.listdir()
if 'input.txt' not in files_at_folder:
    print('Ошибка: в текущей директории нет input.txt, создайте его по шаблону в before-usage/input-data-template.txt')
    exit(1)

if 'regions.txt' not in files_at_folder:
    print('Ошибка: в текущей директории нет regions.txt, создайте его и добавьте интересующий вас регион')
    exit(1)

# Файл, полученный в Google Developer Console
with open('input.txt', 'r', encoding='utf-8') as f:
    spreadsheet_id, from_sheet, row_range, columns, CREDENTIALS_FILE = f.read().split('\n')[:5]
    start, end = map(int, row_range.split(','))
with open('regions.txt', 'r', encoding='utf-8') as f:
    regions = f.read().split('\n')
if '' in regions:
    regions.remove('')
if len(regions) == 0:
    regions.append('')

students_sheet = "Сводка по региону"
school_sheet = "Статистика школ"
region_sheet = "Статистика регионов"

service = connect(CREDENTIALS_FILE)
print('Сервисный аккаунт подключился к документу, начинается процесс построения статистики...')
spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
sheet_list = spreadsheet.get('sheets')

prepare_sheet(students_sheet, sheet_list)
prepare_sheet(school_sheet, sheet_list)
prepare_sheet(region_sheet, sheet_list)

data = read(from_sheet, 'A' + str(start), 'Z' + str(end))['values']
data = [x[:26] + [''] * (26 - len(x)) for x in data]
inds = [ord(x) - ord('A') for x in columns]
all_regions = [x[inds[1]] for x in data if len(x) > inds[1]]
while '' in all_regions:
    all_regions.remove('')
students = select_region(data, regions, inds)
if len(students) > 0:
    print()
    print('На вкладку "{name}" добавлен список олимпиадников в выбранном регионе.'.format(name=students_sheet))
    print('Список отсортирован по количеству учеников в школах.')
    write(students_sheet, 'A1', students, "ROWS")

print()
print('На вкладку "{name}" добавлена статистика по школам.'.format(name=school_sheet))
write(school_sheet, 'A1', get_statistics([x[2] for x in students]), "ROWS")

print()
print('На вкладку "{name}" добавлена статистика по регионам.'.format(name=region_sheet))
write(region_sheet, 'A1', get_statistics(all_regions), "ROWS")