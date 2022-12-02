import yfinance
import openpyxl
import shutil
import pathlib
import os
from datetime import datetime
import scrap_mod


class Securities:
    """Parent class"""

    def __init__(self, security_code):
        self.security_code = security_code


class Stock(Securities):
    """child class"""

    def __init__(self, security_code):
        """ """
        super().__init__(security_code)

        ticker_info = yfinance.Ticker(security_code).info
        self.name = ticker_info['shortName']
        self.price = [ticker_info['currentPrice'], ticker_info['currency']]
        self.exchange = ticker_info['exchange']
        self.shares = ticker_info['sharesOutstanding']
        self.report_currency = ticker_info['financialCurrency']
        self.is_df = scrap_mod.get_income_statement(security_code)
        self.bs_df = scrap_mod.get_balance_sheet(security_code)

    def create_val_xlsx(self):
        """Return a raw_fin_data xlxs for the stock"""

        new_val_name = ""

        # Copy the latest Valuation template
        cwd = pathlib.Path.cwd().resolve()
        try:
            template_folder_path = cwd / 'Template'
            if pathlib.Path(template_folder_path).exists():
                template_file_path = [val_file_path for val_file_path in template_folder_path.iterdir()
                                      if template_folder_path.is_dir() and val_file_path.is_file()]
                if len(template_file_path) > 1 or len(template_file_path) == 0:
                    raise FileNotFoundError("The template file error", "temp_file")
            else:
                raise FileNotFoundError("The template folder doesn't exist", "temp_folder")
        except FileNotFoundError as err:
            if err.args[1] == "temp_folder":
                print("The template folder doesn't exist")
            if err.args[1] == "temp_file":
                print("The template file error")
        else:
            new_val_name = self.security_code + " " + os.path.basename(template_file_path[0])
            if not pathlib.Path(new_val_name).exists():
                shutil.copy(template_file_path[0], new_val_name)

        # load and update the new valuation xlsx
        wb = openpyxl.load_workbook(new_val_name)
        self.update_dashboard(wb)
        self.update_data(wb)

        wb.save(new_val_name)

    def update_dashboard(self, wb):
        """Update the Dashboard sheet"""

        dash_sheet = wb['Dashboard']
        dash_sheet.cell(row=3, column=3).value = self.security_code
        dash_sheet.cell(row=4, column=3).value = self.name
        dash_sheet.cell(row=5, column=3).value = datetime.today().strftime('%Y-%m-%d')
        dash_sheet.cell(row=6, column=3).value = self.exchange
        dash_sheet.cell(row=7, column=3).value = self.price[0]
        dash_sheet.cell(row=7, column=4).value = self.price[1]
        dash_sheet.cell(row=8, column=3).value = self.shares
        dash_sheet.cell(row=13, column=3).value = self.report_currency
        dash_sheet.cell(row=14, column=3).value = scrap_mod.get_forex_rate(self.price[1], self.report_currency)

    def update_data(self, wb):
        """Update the Data sheet"""

        data_sheet = wb['Data']
        data_sheet.cell(row=3, column=3).value = self.is_df.columns[0]  # last financial year
        # figures in
        figures_in = int((len(str(self.is_df.iloc[0, 0])) - 9) / 3 + 0.99) * 1000
        data_sheet.cell(row=4, column=3).value = figures_in
        # load income statement
        for i in range(len(self.is_df.columns)):
            data_sheet.cell(row=7, column=i + 3).value = int(self.is_df.iloc[0, i] / figures_in)
            data_sheet.cell(row=9, column=i + 3).value = int(self.is_df.iloc[1, i] / figures_in)
            data_sheet.cell(row=11, column=i + 3).value = int(self.is_df.iloc[2, i] / figures_in)
            data_sheet.cell(row=17, column=i + 3).value = int(self.is_df.iloc[3, i] / figures_in)
            data_sheet.cell(row=18, column=i + 3).value = int(self.is_df.iloc[4, i] / figures_in)
        # load balance sheet
        for i in range(1, len(self.bs_df.columns)):
            data_sheet.cell(row=20, column=i + 3).value = int(self.bs_df.iloc[0, i] / figures_in)
            data_sheet.cell(row=21, column=i + 3).value = int(self.bs_df.iloc[1, i] / figures_in)
            data_sheet.cell(row=22, column=i + 3).value = int(self.bs_df.iloc[2, i] / figures_in)
            data_sheet.cell(row=23, column=i + 3).value = int(self.bs_df.iloc[3, i] / figures_in)
            data_sheet.cell(row=25, column=i + 3).value = int(self.bs_df.iloc[4, i] / figures_in)
            data_sheet.cell(row=26, column=i + 3).value = int(self.bs_df.iloc[5, i] / figures_in)
            data_sheet.cell(row=27, column=i + 3).value = int(self.bs_df.iloc[6, i] / figures_in)
            data_sheet.cell(row=28, column=i + 3).value = int(self.bs_df.iloc[7, i] / figures_in)

    def export_statements(self):
        """Export the income statement and balance sheet"""

        self.is_df.to_csv(f'{self.security_code}_income_statement.csv', sep=',', encoding='utf-8')
        self.bs_df.to_csv(f'{self.security_code}_balance_sheet.csv', sep=',', encoding='utf-8')
