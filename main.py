from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Font, Border, NamedStyle
from openpyxl.formatting.rule import CellIsRule


wb = load_workbook("Analyze_Report_20210914_125620.xlsx")
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

#  format cell col D-F - Number
for row in ws1.iter_rows(min_row=2, min_col=4, max_col=6):
    for cell in row:
        cell.number_format = '0.00'

#  col A-F format as table
tab = Table(displayName=ws1.title+"_table", ref="A1:" + get_column_letter(ws1.max_column) + str(ws1.max_row))
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=False, showColumnStripes=False)
tab.tableStyleInfo = style

ws1.add_table(tab)

#  col D-F (minus headers) conditional formatting
redFill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
redFont = Font(color="9C0006")
yellowFill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
yellowFont = Font(color="9C5700")
greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
greenFont = Font(color="006100")

ws1.conditional_formatting.add("D2:D" + str(ws1.max_row),
                              CellIsRule(operator='between', formula=['0.1', '0.4'], stopIfTrue=True, fill=redFill, font=redFont))
ws1.conditional_formatting.add("D2:D" + str(ws1.max_row),
                              CellIsRule(operator='between', formula=['0.4', '0.7'], stopIfTrue=True, fill=yellowFill, font=yellowFont))
ws1.conditional_formatting.add("D2:D" + str(ws1.max_row),
                              CellIsRule(operator='greaterThan', formula=['0.7'], stopIfTrue=True, fill=greenFill, font=greenFont))

#  	>=0.7 light green
#  	>=0.4 and <0.7 yellow
#  	>0 and <0.4 red

wb.save("py_Analyze_Report_20210914_125620.xlsx")