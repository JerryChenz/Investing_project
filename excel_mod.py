import openpyxl
import scrap_mod

def get_raw_data(ticker):
    """Return a raw_fin_data xlxs for the ticker"""

    # Initialize the stock object given ticker



def write_excel(filename):

    wb = openpyxl.load_workbook(filename)
    sheet = wb['Sheet1']
    sheet.cell(row=1, column=2).value = 'test'
    wb.save('Raw_fin_data.xlsx')


