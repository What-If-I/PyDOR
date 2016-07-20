import os


def search_in_column(sheet, value, search_in_row, start, end):
    """
    Находит ближайшую ячейки по указанному содержимому.
    :param sheet: Название листа
    :param value: Искомое значение
    :param search_in_row: Поиск в строке
    :param start: Колонка начало поиска
    :param end: Колонка окончания поиска
    :return Координаты ячейки:
    """
    for column in range(start, end + 1):
        active_column = sheet.cell(row=search_in_row, column=column)
        if active_column.value == value:
            return active_column
    return False


def search_in_row(sheet, value, search_in_column, start, end):
    """
    Находит ближайшую ячейки по указанному содержимому.
    :param sheet: Название листа
    :param value: Искомое значение
    :param search_in_column: Поиск в колонке
    :param start: Колонка начало поиска
    :param end: Колонка окончания поиска
    :return Координаты ячейки:
    """
    for row in range(start, end + 1):
        active_column = sheet.cell(row=row, column=search_in_column)
        if active_column.value == value:
            return active_column
    return False


def reports_name_and_path(exclude_folder=''):
    for dir_path, dir_names, file_names in os.walk(os.getcwd()):
        if exclude_folder in dir_names:
            dir_names.remove(exclude_folder)
        for file_name in file_names:
            yield file_name, dir_path


def find_report(reports, beginning):
    for report in reports:
        if report[0].startswith(beginning):
            report_lst = os.path.join(report[1], report[0])
            return report_lst
    return None


def get_sec(s):
    times = s.split(':')
    seconds = 0
    i = 0
    for time in reversed(times):
        seconds += int(time) * 60 ** i
        i += 1
    return seconds
