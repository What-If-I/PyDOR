import openpyxl
import datetime
from dor_pars import *
from openpyxl.cell import get_column_letter, column_index_from_string

directory = '../Daily Daily reports'
os.chdir(directory)

# Сегодняшний день и месяц
today = datetime.date.today()
current_month = "{:%B}".format(today)
current_day = today.day

# отркываем файл DOR
dor = openpyxl.load_workbook("DOR.xlsx")

# Список всех файлов в директории
reports = [(a, b) for a, b in reports_name_and_path(exclude_folder="_old")]

# ================ Начало АА и АА Сервис =============================

# открываем страницу AA, находим столбец текущего дня
aa_dor_sheet = dor.get_sheet_by_name("AA")
curr_month_cell = search_in_column(aa_dor_sheet, current_month, 1,
                                   start=1, end=aa_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(aa_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=aa_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами по АА сервис
aa_sc_reports = find_report(reports, "AA-Service-Centre_")
aa_reports = find_report(reports, "AA_")

# открываем отчёт AA-SC
aa_sc = openpyxl.load_workbook(aa_sc_reports[0])
aa_sc_sheet = aa_sc.active
aa_sc_statistic = {"entered": aa_sc_sheet['D9'].value,
                   "answered": aa_sc_sheet['E9'].value,
                   "answered<sl": aa_sc_sheet['G9'].value,
                   "abandoned": aa_sc_sheet['F9'].value
                   }

# Заполняем DOR AA Service Center
aa_dor_sheet.cell(column=cur_day_column_index, row=15).value = aa_sc_statistic["entered"]
aa_dor_sheet.cell(column=cur_day_column_index, row=16).value = aa_sc_statistic["answered"]
aa_dor_sheet.cell(column=cur_day_column_index, row=17).value = aa_sc_statistic["answered<sl"]
aa_dor_sheet.cell(column=cur_day_column_index, row=18).value = aa_sc_statistic["abandoned"]

# открываем отчёт AA
aa_report = openpyxl.load_workbook(aa_reports[0])
aa_report_sheet = aa_report.active
aa_statistic = {"entered": aa_report_sheet['D9'].value,
                "answered": aa_report_sheet['E9'].value,
                "answered<sl": aa_report_sheet['G9'].value,
                "abandoned": aa_report_sheet['F9'].value,
                "AHT": get_sec(aa_report_sheet["P9"].value)
                }

# Заполняем DOR AA
aa_dor_sheet.cell(column=cur_day_column_index, row=5).value = aa_statistic["AHT"]
aa_dor_sheet.cell(column=cur_day_column_index, row=7).value = aa_statistic["entered"]
aa_dor_sheet.cell(column=cur_day_column_index, row=8).value = aa_statistic["answered"]
aa_dor_sheet.cell(column=cur_day_column_index, row=9).value = aa_statistic["answered<sl"]
aa_dor_sheet.cell(column=cur_day_column_index, row=10).value = aa_statistic["abandoned"]

# ============================================================================================

# ======================= Начало BSC =========================================================

# открываем страницу BSC, находим столбец текущего дня
bsc_dor_sheet = dor.get_sheet_by_name("BSC")
curr_month_cell = search_in_column(bsc_dor_sheet, current_month, 1, start=1, end=bsc_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(bsc_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=bsc_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
bsc_reports = find_report(reports, "BSC_")

# открываем отчёт BSC
bsc_report_wb = openpyxl.load_workbook(bsc_reports[0])
bsc_report_sheet = bsc_report_wb.active
bsc_statistic = {"entered": bsc_report_sheet['D9'].value,
                 "answered": bsc_report_sheet['E9'].value,
                 "answered<sl": bsc_report_sheet['G9'].value,
                 "abandoned": bsc_report_sheet['F9'].value,
                 "ghost_calls": bsc_report_sheet['J9'].value,
                 "AHT": get_sec(bsc_report_sheet['P9'].value)
                 }

# Заполняем DOR BSC

bsc_dor_sheet.cell(column=cur_day_column_index, row=5).value = bsc_statistic["AHT"]
bsc_dor_sheet.cell(column=cur_day_column_index, row=7).value = bsc_statistic["entered"]
bsc_dor_sheet.cell(column=cur_day_column_index, row=8).value = bsc_statistic["answered"]
bsc_dor_sheet.cell(column=cur_day_column_index, row=9).value = bsc_statistic["answered<sl"]
bsc_dor_sheet.cell(column=cur_day_column_index, row=10).value = bsc_statistic["abandoned"]
bsc_dor_sheet.cell(column=cur_day_column_index, row=11).value = bsc_statistic["ghost_calls"]

# ===================================================================================================

# ============================ Buderus ===============================================================

# открываем страницу Buderus, находим столбец текущего дня
buderus_dor_sheet = dor.get_sheet_by_name("Buderus")
curr_month_cell = search_in_column(buderus_dor_sheet, current_month, 1, start=1, end=buderus_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(buderus_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=buderus_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
buderus_reports = find_report(reports, "Buderus_")

# открываем отчёт Buderus
buderus_report_wb = openpyxl.load_workbook(buderus_reports[0])
buderus_report_sheet = buderus_report_wb.active
buderus_statistic = {"entered": buderus_report_sheet['D9'].value,
                     "answered": buderus_report_sheet['E9'].value,
                     "answered<sl": buderus_report_sheet['G9'].value,
                     "abandoned": buderus_report_sheet['F9'].value,
                     "AHT": get_sec(buderus_report_sheet['P9'].value)
                     }

# Заполняем DOR Buderus

buderus_dor_sheet.cell(column=cur_day_column_index, row=5).value = buderus_statistic["AHT"]
buderus_dor_sheet.cell(column=cur_day_column_index, row=7).value = buderus_statistic["entered"]
buderus_dor_sheet.cell(column=cur_day_column_index, row=8).value = buderus_statistic["answered"]
buderus_dor_sheet.cell(column=cur_day_column_index, row=9).value = buderus_statistic["answered<sl"]
buderus_dor_sheet.cell(column=cur_day_column_index, row=10).value = buderus_statistic["abandoned"]

# ================================================================================================

# ============================ Buderus-Sales ===============================================================

# открываем страницу Buderus-Sales, находим столбец текущего дня
buderus_sales_dor_sheet = dor.get_sheet_by_name("Buderus Sale")
curr_month_cell = search_in_column(buderus_sales_dor_sheet, current_month, 1,
                                   start=1, end=buderus_sales_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(buderus_sales_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=buderus_sales_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
buderus_sales_reports = find_report(reports, "Buderus-Sale_")

# открываем отчёт Buderus-Sales
buderus_sales_report_wb = openpyxl.load_workbook(buderus_sales_reports[0])
buderus_sales_report_sheet = buderus_sales_report_wb.active
buderus_sales_statistic = {"entered": buderus_sales_report_sheet['D9'].value,
                           "answered": buderus_sales_report_sheet['E9'].value,
                           "answered<sl": buderus_sales_report_sheet['G9'].value,
                           "abandoned": buderus_sales_report_sheet['F9'].value,
                           "AHT": get_sec(buderus_sales_report_sheet['P9'].value)
                           }

# Заполняем DOR Buderus-Sales

buderus_sales_dor_sheet.cell(column=cur_day_column_index, row=5).value = buderus_sales_statistic["AHT"]
buderus_sales_dor_sheet.cell(column=cur_day_column_index, row=7).value = buderus_sales_statistic["entered"]
buderus_sales_dor_sheet.cell(column=cur_day_column_index, row=8).value = buderus_sales_statistic["answered"]
buderus_sales_dor_sheet.cell(column=cur_day_column_index, row=9).value = buderus_sales_statistic["answered<sl"]
buderus_sales_dor_sheet.cell(column=cur_day_column_index, row=10).value = buderus_sales_statistic["abandoned"]

# ================================================================================================

# ============================ Michelin ===============================================================

# открываем страницу Michelin, находим столбец текущего дня
michelin_dor_sheet = dor.get_sheet_by_name("Michelin")
curr_month_cell = search_in_column(michelin_dor_sheet, current_month, 1,
                                   start=1, end=michelin_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(michelin_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=michelin_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
michelin_calls_reports = find_report(reports, "Michelin-Calls_")
michelin_votes_reports = find_report(reports, "Michelin-Votes_")

# открываем отчёт Michelin-calls и собираем статистику
michelin_calls_report_wb = openpyxl.load_workbook(michelin_calls_reports)
michelin_calls_report_sheet = michelin_calls_report_wb.active
michelin_calls_statistic = {"entered": michelin_calls_report_sheet['C6'].value,
                            "answered": michelin_calls_report_sheet['D6'].value,
                            "answered<sl": michelin_calls_report_sheet['E6'].value,
                            "abandoned": michelin_calls_report_sheet['H6'].value,
                            "abandoned<5": michelin_calls_report_sheet['I6'].value,
                            "AHT": get_sec(michelin_calls_report_sheet['M6'].value),
                            }

# открываем отчёт Michelin-votes и собираем статистику
michelin_votes_report_wb = openpyxl.load_workbook(michelin_votes_reports)
michelin_votes_report_sheet = michelin_votes_report_wb.get_sheet_by_name("Kazan Michelin Voting Total")
michelin_votes_statistic = {"vote1": michelin_votes_report_sheet['B5'].value,
                            "vote2": michelin_votes_report_sheet['C5'].value,
                            "vote3": michelin_votes_report_sheet['D5'].value,
                            "vote4": michelin_votes_report_sheet['E5'].value,
                            "vote5": michelin_votes_report_sheet['F5'].value
                            }

# Заполняем DOR Michelin

michelin_dor_sheet.cell(column=cur_day_column_index, row=5).value = michelin_calls_statistic["AHT"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=7).value = michelin_calls_statistic["entered"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=8).value = michelin_calls_statistic["answered"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=9).value = michelin_calls_statistic["answered<sl"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=10).value = michelin_calls_statistic["abandoned"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=11).value = michelin_calls_statistic["abandoned<5"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=12).value = michelin_votes_statistic["vote1"] + \
                                                                     michelin_votes_statistic["vote2"]
michelin_dor_sheet.cell(column=cur_day_column_index, row=13).value = sum(michelin_votes_statistic.values())

# ================================================================================================

# ============================ Invitro ===============================================================

# открываем страницу Invitro, находим столбец текущего дня
invitro_dor_sheet = dor.get_sheet_by_name("Invitro")
curr_month_cell = search_in_column(invitro_dor_sheet, current_month, 1,
                                   start=1, end=invitro_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(invitro_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=invitro_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
invitro_calls_reports = find_report(reports, "Invitro_Calls(Cities+Expert)_")
invitro_status_reports = find_report(reports, "Invitro-Status_")

# открываем отчёт Invitro-calls и собираем статистику
invitro_calls_report_wb = openpyxl.load_workbook(invitro_calls_reports)
invitro_calls_report_sheet = invitro_calls_report_wb.active
invitro_calls_statistic = {"entered": invitro_calls_report_sheet['D6'].value,
                           "answered": invitro_calls_report_sheet['E6'].value,
                           "answered<sl": invitro_calls_report_sheet['F6'].value,
                           "abandoned": invitro_calls_report_sheet['G6'].value,
                           "abandoned<5": invitro_calls_report_sheet['H6'].value,
                           "AHT": get_sec(invitro_calls_report_sheet['M6'].value),
                           "ATT": get_sec(invitro_calls_report_sheet['L6'].value),
                           "ACW": get_sec(invitro_calls_report_sheet['Q6'].value),
                           }

# открываем отчёт Invitro-status и собираем статистику
invitro_status_report_wb = openpyxl.load_workbook(invitro_status_reports[0])
invitro_status_report_sheet = invitro_status_report_wb.active

invitro_status_total_row = search_in_row(invitro_status_report_sheet, "Total:", 1,
                                         start=1, end=invitro_status_report_sheet.max_row).row

# invitro_status_total_cell_coordinates = {"on_call": "E{}".format(invitro_status_total_row),
#                                          "after_call": "F{}".format(invitro_status_total_row),
#                                          "after_call_manual": "G{}".format(invitro_status_total_row),
#                                          "chat": "K{}".format(invitro_status_total_row),
#                                          "back_office_work": "J{}".format(invitro_status_total_row),
#                                          "available": "D{}".format(invitro_status_total_row)
#                                          }
#
#
# invitro_status_statistic = {"on_call": invitro_status_report_sheet["E4"].value,
#                             "after_call": invitro_status_report_sheet['F4'].value,
#                             "after_call_manual": invitro_status_report_sheet['G4'].value,
#                             "chat": invitro_status_report_sheet['K4'].value,
#                             "back_office_work": invitro_status_report_sheet['F5'].value,
#                             "available": invitro_status_report_sheet['F5'].value
#                             }

# Заполняем DOR Invitro

invitro_dor_sheet.cell(column=cur_day_column_index, row=5).value = invitro_calls_statistic["AHT"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=6).value = invitro_calls_statistic["ATT"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=7).value = invitro_calls_statistic["ACW"]

invitro_dor_sheet.cell(column=cur_day_column_index, row=10).value = invitro_calls_statistic["entered"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=11).value = invitro_calls_statistic["answered"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=12).value = invitro_calls_statistic["answered<sl"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=13).value = invitro_calls_statistic["abandoned"]
invitro_dor_sheet.cell(column=cur_day_column_index, row=14).value = invitro_calls_statistic["abandoned<5"]

# invitro_dor_sheet.cell(column=cur_day_column_index, row=12).value = invitro_status_votes_statistic["vote1"]

# ================================================================================================


# ============================ Invitro-Expert ===============================================================

# открываем страницу Invitro-Expert, находим столбец текущего дня
invitro_expert_dor_sheet = dor.get_sheet_by_name("Expert")
curr_month_cell = search_in_column(invitro_expert_dor_sheet, current_month, 1,
                                   start=1, end=invitro_expert_dor_sheet.max_column)
curr_month_column_index = column_index_from_string(curr_month_cell.column)
curr_day_cell = search_in_column(invitro_expert_dor_sheet, current_day, 2,
                                 start=curr_month_column_index, end=invitro_expert_dor_sheet.max_column)
cur_day_column_index = column_index_from_string(curr_day_cell.column)

# находим файл с отчётами
invitro_expert_calls_reports = find_report(reports, "Expert-Calls_")

# открываем отчёт Invitro-Expert-calls и собираем статистику
invitro_expert_calls_report_wb = openpyxl.load_workbook(invitro_expert_calls_reports)
invitro_expert_calls_report_sheet = invitro_expert_calls_report_wb.active
invitro_epxert_calls_statistic = {"entered": invitro_expert_calls_report_sheet['D6'].value,
                                  "answered": invitro_expert_calls_report_sheet['E6'].value,
                                  "answered<sl": invitro_expert_calls_report_sheet['F6'].value,
                                  "abandoned": invitro_expert_calls_report_sheet['G6'].value,
                                  "abandoned<5": invitro_expert_calls_report_sheet['H6'].value,
                                  "AHT": get_sec(invitro_expert_calls_report_sheet['M6'].value),
                                  "ATT": get_sec(invitro_expert_calls_report_sheet['L6'].value),
                                  "ACW": get_sec(invitro_expert_calls_report_sheet['Q6'].value),
                                  }

# Заполняем DOR Invitro-Expert

invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=5).value = invitro_epxert_calls_statistic["AHT"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=6).value = invitro_epxert_calls_statistic["ATT"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=7).value = invitro_epxert_calls_statistic["ACW"]

invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=10).value = invitro_epxert_calls_statistic["entered"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=11).value = invitro_epxert_calls_statistic["answered"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=12).value = invitro_epxert_calls_statistic["answered<sl"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=13).value = invitro_epxert_calls_statistic["abandoned"]
invitro_expert_dor_sheet.cell(column=cur_day_column_index, row=14).value = invitro_epxert_calls_statistic["abandoned<5"]

# ================================================================================================

dor.save("DOR.xlsx")
print("Done!")
