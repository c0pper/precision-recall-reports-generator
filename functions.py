from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import CellIsRule
from openpyxl import load_workbook
import glob

def apply_conditional_format(worksheet, range:str, redmin:str, redmax:str, yellowmin:str, yellowmax:str, greenmin:str):
    redFill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    redFont = Font(color="9C0006")
    yellowFill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    yellowFont = Font(color="9C5700")
    greenFill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    greenFont = Font(color="006100")

    #  	>0 and <0.4 red
    worksheet.conditional_formatting.add(range,
                                  CellIsRule(operator='between', formula=[redmin, redmax], stopIfTrue=True, fill=redFill, font=redFont))

    #  	>=0.4 and <0.7 yellow
    worksheet.conditional_formatting.add(range,
                                  CellIsRule(operator='between', formula=[yellowmin, yellowmax], stopIfTrue=True, fill=yellowFill, font=yellowFont))

    #  	>=0.7 light green
    worksheet.conditional_formatting.add(range,
                                  CellIsRule(operator='greaterThan', formula=[greenmin], stopIfTrue=True, fill=greenFill, font=greenFont))


#  get 2nd most recent precision and recall
def get_2nd_most_recent_report():
    import glob
    import os

    files = glob.glob("raw_reports/formatted/*.xlsx")
    files.sort(key=os.path.getmtime)
    wb = load_workbook(files[-2])

    return wb

#  most recent precision and recall - (subtract) 2nd most recent precision and recall
def precision_recall_difference():
    previous_wb = get_2nd_most_recent_report()
    previous_wb_worksheet = previous_wb.worksheets[0]
    main_wb = load_workbook("Report.xlsx")

    #  check if workbook name in previous wb is in main report
    ws_title = previous_wb.active.title
    if ws_title in main_wb.sheetnames:
        print("sheet already done")
    else:
        main_wb.create_sheet(ws_title)

        # calculate total number of rows and columns in source excel file
        mr = previous_wb_worksheet.max_row
        mc = previous_wb_worksheet.max_column

        # copying the cell values from source
        # excel file to destination excel file
        for i in range(1, mr + 1):
            for j in range(1, mc + 1):
                # reading cell value from source excel file
                c = previous_wb_worksheet.cell(row=i, column=j)

                # writing the read value to destination excel file
                main_wb[ws_title].cell(row=i, column=j).value = c.value

        main_wb.save("Report.xlsx")


    #  https://www.geeksforgeeks.org/python-how-to-copy-data-from-one-excel-sheet-to-another/


def find_already_formatted_files():
    already_formatted_files = []
    for file in glob.glob("formatted/*.xlsx"):
        file = file.split("\\")[-1]
        already_formatted_files.append(file)
    return already_formatted_files




if __name__ == "__main__":
    precision_recall_difference()