import win32com.client
import os
import re
import psycopg2
from datetime import datetime, timedelta
from time import sleep
from shutil import copyfile
from contextlib import closing
from psycopg2.extras import DictCursor


class Report():
    """Класс -Отчет- для автоматического обновления \n
        Методы: \n
                open - открытие отчета в excel'e
                refresh - обновление данных в отчете
                save - сохранение отчета с текущей датой в имени
                copy - копирование отчета по copy_path
                close - закрытие excel
                delete - удаление отчета по возможности
                update - полный цикл обновления отчета
                update_without_copy - цикл обновления отчета без последующего копирования"""

    def __init__(self, name, report_path, copy_path='', previous_date='', current_date=''):
        """Инициализация атрибутов отчета"""
        self.open_name = name + ' ' + previous_date + '.xlsx'
        self.save_name = name + ' ' + current_date + '.xlsx'
        self.open_report = report_path + self.open_name
        self.save_report = report_path + self.save_name
        self.copy_report = copy_path + self.save_name

        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        self.excel.Visible = 0
        self.work_book = None
        print(f'Инициализация отчета {name} выполнена')

    def open(self):
        """Открытие отчета в рабочей книге"""
        print(f'Открывается отчет: {self.open_report}')
        self.work_book = self.excel.Workbooks.Open(self.open_report)
        print('Открытие завершено')
        sleep(3)

    def refresh(self):
        """Обновление всех данных в отчете. Ожидание пока не завершится обновление"""
        print(f'Обновление отчета: {self.open_name}')
        self.work_book.RefreshAll()
        self.excel.CalculateUntilAsyncQueriesDone()
        print('Обновление завершено')
        sleep(3)

    def save(self):
        """Сохранение отчета с текущей датой в имени"""
        print(f'Сохранение отчета: {self.save_report}')
        self.work_book.SaveAs(self.save_report)
        print('Сохранение завершено')
        sleep(3)

    def copy(self):
        """Копирование отчета по указанному пути"""
        try:
            print(f'Копирование отчета: {self.copy_report}')
            copyfile(self.save_report, self.copy_report)
            print('Копирование завершено')
        except:
            print('Ошибка при копировании')
        sleep(3)

    def close(self):
        """Закрытие excel"""
        self.excel.Quit()
        print('Отчет закрыт')
        sleep(3)

    def delete(self):
        """Удаление отчета по возможности"""
        try:
            os.remove(self.open_report)
            print(f'Удален: {self.open_report}')
        except:
            print(f'Ошибка при удалении отчета: {self.open_report}')

    def update(self):
        """Полный цикл обновления отчета"""
        self.open()
        self.refresh()
        self.refresh()
        self.save()
        self.close()
        self.delete()
        self.copy()
        print('Цикл обновления отчета завершен')

    def update_without_copy(self):
        """Цикл обновления отчета без последующего копирования"""
        self.open()
        self.refresh()
        self.refresh()
        self.save()
        self.close()
        self.delete()
        print('Цикл обновления отчета завершен')


def get_dates():
    """Словарь с датами"""
    date_format = '%d.%m'
    update_format = '%Y-%m-%d %H:%m:%S'
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    week_ago = today - timedelta(days=7)
    return {'today': today.strftime(date_format),
            'yesterday': yesterday.strftime(date_format),
            'week_ago': week_ago.strftime(date_format),
            'update_date': today.strftime(update_format)}


def get_report_list(folder_path):
    """Список уникальных имен отчетов в папке"""
    all_reports = list(filter(lambda x: x.endswith(
        '.xlsx'), os.listdir(path=folder_path)))
    report_list = []
    for r in all_reports:
        report_list.append(' '.join(re.findall(r'[А-Яа-я]+', r)))
    return set(report_list)


# Проверяем выполнение процедур в БД по дате
with closing(psycopg2.connect(dbname='analytics', user='', password='', host='', port='5432')) as conn:
    with conn.cursor(cursor_factory=DictCursor) as cursor:
        cursor.execute(
            "select to_char(modified, 'dd.mm') from public.tables_properties where service='cargo_handling_daily' and property='last_update'")
        reports_update = cursor.fetchone()

while reports_update[0] != get_dates()['today']:
    print(f'На {get_dates()["update_date"]} данные для отчетов не обновлены')
    sleep(900)
    with closing(psycopg2.connect(dbname='analytics', user='', password='', host='', port='5432')) as conn:
        with conn.cursor(cursor_factory=DictCursor) as cursor:
            cursor.execute(
                "select to_char(modified, 'dd.mm') from public.tables_properties where service='cargo_handling_daily' and property='last_update'")
            reports_update = cursor.fetchone()

print('Данные для отчетов подготовлены, начато обновление')
print('==================================================')

# Пути до папок с отчетами
report_path = r"\\10-fs03\\Users\\Moscow\\Proekt Punkti Vidachi\\ОСТАТКИ ТЕРМИНАЛОВ\\"
copy_path = r"\\10-fs03\\Users\\Yekaterinburg\\Департамент аналитики и развития бизнес-процессов\\Группа анализа\\ОТЧЕТЫ\\ОТЧЕТ_ОСТАТКИ\\Архив\\"

# Обновление ежедневных отчетов
for name in get_report_list(report_path):
    report = Report(name, report_path, copy_path, get_dates()
                    ['yesterday'], get_dates()['today'])
    if name in ['Отчет по задержкам']:
        report.update_without_copy()
    else:
        report.update()

# Обновление еженедельных отчетов
if datetime.today().isoweekday() == 2:
    report_path = r"\\10-fs03\\Users\\Moscow\\Proekt Punkti Vidachi\\Отчеты для региональной сети\\Соблюдение сроков доставки\\Соблюдение сроков обработки на терминалах сети\\"
    for name in get_report_list(report_path):
        report = Report(name, report_path, copy_path, get_dates()[
                        'week_ago'], get_dates()['today'])
        report.update_without_copy()

# Ожидание закрытия программы
print()
print('Выполнение программы завершено')
sleep(3600)
