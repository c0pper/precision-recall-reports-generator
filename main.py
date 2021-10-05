from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from functions import apply_conditional_format, find_already_formatted_files
import glob, os

os.chdir("raw_reports")


def format_report(file):
    wb = load_workbook(file)
    ws1 = wb.active

    #  delete col D-M
    ws1.delete_cols(4, 10)

    #  delete col G-I
    ws1.delete_cols(7, 9)

    #  enlarge col A
    for cell in ws1["A"]:
        length = max(len(cell.value) for cell in ws1["A"])
        ws1.column_dimensions["A"].width = length

    #  enlarge col Precions, F Measure
    ws1.column_dimensions["D"].width = 12
    ws1.column_dimensions["F"].width = 12

    #  format cell col D-J - Number
    for row in ws1.iter_rows(min_row=2, min_col=4, max_col=10):
        for cell in row:
            cell.number_format = '0.00'

    #  col A-F format as table
    print(ws1.title+"_table")
    tab = Table(displayName=file+"_table", ref="A1:" + "H" + str(ws1.max_row))
    style = TableStyleInfo(name="TableStyleLight13", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=False, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws1.add_table(tab)

    #  col D-F (minus headers) conditional formatting
    apply_conditional_format(ws1, "D2:F" + str(ws1.max_row), "0.1", "0.39", "0.4", "0.69", "0.7")

    #  Differential with previous report
    ws1["G1"] = "Precision Diff"
    ws1["H1"] = "Recall Diff"

    #  add precision and recall average in col J-K
    ws1.column_dimensions["I"].width = 16

    ws1["I2"] = "Precision Average"
    ws1["J2"] = f"=AVERAGE(D2:D{str(ws1.max_row)})"

    ws1["I3"] = "Recall Average"
    ws1["J3"] = f"=AVERAGE(E2:E{str(ws1.max_row)})"

    wb.save(f"formatted/{file}")


if __name__ == "__main__":
    raw_reports = glob.glob("*.xlsx")
    for file in raw_reports:
        if file not in find_already_formatted_files():
            format_report(file)
        else:
            print(f"{file} already formatted")
