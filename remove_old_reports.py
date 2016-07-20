import os
import datetime
from shutil import copyfile
from os.path import join


def main():
    path = './Daily reports'  # Путь к папкам с отчётами
    if os.access(path, 1):  # Проверяем наличие доступа к папке
        os.chdir(path)
    else:
        input("Access denied or 'Daily reports' folder may not exist.")
        exit()

    reports_path = os.getcwd()  # Запомнили полный путь к директории с отчётами
    old_reports_folder = join(reports_path, '_old')  # Путь к папке для бэкапа отчётов

    if not os.access(old_reports_folder, 1):
        os.mkdir(old_reports_folder)

    # Получаем список папок в директории ./Daily reports. Все файлы исключаются.
    folders = [f for f in os.listdir(reports_path) if
               os.path.isdir(join(reports_path, f)) and
               f != old_reports_folder
               ]
    reports = {}  # Словарь для путей к отчётам и их названиям. { Название отчёта : путь к нему} 

    # Проходимся по каждой папке с отчётами и заносим {название отчёта : путь} в словарь
    for folder in folders:
        reports_folder = join(os.getcwd(), folder)
        files = [f for f in next(os.walk(reports_folder))[2]]
        for file in files:
            reports[file] = join(reports_folder)

    # Новый словарь с датами создания отчётов {название отчёта: дата в формате Y-m-d
    report_dates = {}
    for key in reports:
        report_dates[key] = datetime.datetime.strptime(key[-15:-25:-1][::-1], "%Y-%m-%d").date()

    # Список с названиями отчётов что старше сегодняшнего дня.
    outdated_reports_lts = [name for name, date in report_dates.items() if date < datetime.date.today()]

    # Словарь содержащий список старых отчётов { имя отчёта : полный путь }
    print("Старые отчёты будут скопированы в папку %s и удалены." % old_reports_folder)
    print("\nСписок старых отчётов\n======================================")
    outdated_reports = {}
    for name, path in reports.items():
        if name in outdated_reports_lts:
            outdated_reports[name] = path
            print(join(path, name))  # Выводим старые отчёты

    print("\nФайлы удалены\n======================================")

    # Копируем каждый отчёт в собственную папку в папке ./Daily reports/ old_reports_folder
    for key, value in outdated_reports.items():
        report_folder = os.path.basename(os.path.normpath(value))
        save_dir = join(old_reports_folder, report_folder)
        if not os.access(save_dir, 1):  # Если папки нет, то мы её создаем.
            os.mkdir(save_dir)
            print("%s folder did not exist and was created." % save_dir)
        copyfile(join(value, key), join(save_dir, key))  # Копируем отчёты

    # Удаляем все старые отчёты
    for name, path in outdated_reports.items():
        os.remove(join(path, name))
        print("%s successfully removed." % join(path, name))

    print("\n%i files were removed." % len(outdated_reports))

if __name__ == '__main__':
    main()
