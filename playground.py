import openpyxl
from dor_pars import *

errors = 0
directory = './Daily reports/'
os.chdir(directory)

# Список всех файлов в директории
reports = [(name, path) for name, path in reports_name_and_path(exclude_folder="_old")]

# Сегодняшний день и месяц
today = datetime.date.today()
current_month = "{:%B}".format(today)
yesterday = today - datetime.timedelta(4)
yesterday_day = yesterday.day

bsc_status_reports = find_report(reports, "BSC-Status_", yesterday)

bsc_status_report_wb = openpyxl.load_workbook(bsc_status_reports)
bsc_status_report_sheet = bsc_status_report_wb.active

bsc_status_total_row = search_in_row(bsc_status_report_sheet, "Total:", 1,
                                     start=1, end=bsc_status_report_sheet.max_row).row

bsc_status_total_cell_coordinates = {"on_call": "E{}".format(bsc_status_total_row),
                                     "after_call": "G{}".format(bsc_status_total_row),
                                     "mail_flex": "H{}".format(bsc_status_total_row),
                                     "back_office_work": "J{}".format(bsc_status_total_row),
                                     "available": "D{}".format(bsc_status_total_row)
                                     }

status_times = get_status_total(bsc_status_report_sheet)
print(bsc_status_report_sheet.merged_cells)
print(bsc_status_report_sheet.merged_cell_ranges)
print(status_times)
