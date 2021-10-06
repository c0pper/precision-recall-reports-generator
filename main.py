from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from functions import check_new_raw_reports, format_sheet, calculate_precision_recall_difference


if __name__ == "__main__":
    check_new_raw_reports()
    calculate_precision_recall_difference()