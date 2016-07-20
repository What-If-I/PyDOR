import os

import openpyxl
import datetime

from locale import setlocale, getlocale, LC_ALL
from openpyxl.cell import get_column_letter, column_index_from_string

# def search_report():
#     walk = os.walk(os.getcwd())
#     print(walk)
#     return None

dirnames = './Daily reports'
os.chdir(dirnames)


def reports_name_and_path(exclude_folder=''):
    for dir_path, dir_names, file_names in os.walk(os.getcwd()):
        if exclude_folder in dir_names:
            dir_names.remove(exclude_folder)
        for file_name in file_names:
            yield file_name, dir_path


def find_report(reports, begining):
    report_lst = []
    for report in reports:
        if report[0].startswith(begining):
            report_lst.append(report)
    return report_lst


reports = []
for f_name, f_path in reports_name_and_path('_old'):
    reports.append((f_name, f_path))

aa_reports = find_report(reports, "AA_")

print(aa_reports)
