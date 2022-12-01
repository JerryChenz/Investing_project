import openpyxl



def write_excel(filename):

    wb = openpyxl.load_workbook(filename)
    sheet = wb['Sheet1']
    sheet.cell(row=1, column=2).value = 'test'
    wb.save('Raw_fin_data.xlsx')


