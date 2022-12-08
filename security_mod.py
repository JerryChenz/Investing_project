import yfinance
import xlwings
import pathlib
import shutil
import os
from datetime import datetime
import pandas as pd
import scrap_mod
import re


class Asset:
    """Parent class"""

    def __init__(self, security_code):
        self.security_code = security_code
        self.name = None
        self.price = None
        self.price_currency = None
        self.exchange = None
        self.ideal_price = None
        self.current_irr = None
        self.risk_premium = None
        self.val_status = None
        self.periodic_payment = None
        self.next_earnings = None  # next coupon date for bonds
        self.invest_horizon = None
        self.unit_cost = None
        self.total_units = None


class Stock(Asset):
    """child class"""

    def __init__(self, security_code):
        """ """
        super().__init__(security_code)

        self.shares = None
        self.report_currency = None
        self.is_df = None
        self.bs_df = None

    def load_from_yf(self):
        """Scrap the data from yahoo finance"""

        ticker_data = yfinance.Ticker(self.security_code)

        self.name = ticker_data.info['shortName']
        self.price = [ticker_data.info['currentPrice'], ticker_data.info['currency']]
        self.exchange = ticker_data.info['exchange']
        self.shares = ticker_data.info['sharesOutstanding']
        self.report_currency = ticker_data.info['financialCurrency']
        self.is_df = scrap_mod.get_income_statement(self.security_code)
        self.bs_df = scrap_mod.get_balance_sheet(self.security_code)
        self.next_earnings = pd.to_datetime(datetime.fromtimestamp(ticker_data.info['mostRecentQuarter'])
                                            .strftime("%Y-%m-%d")) + pd.DateOffset(months=6)

    def create_val_xlsx(self):
        """Return a raw_fin_data xlsx for the stock"""

        new_bool = False
        r = re.compile(".*Valuation_v")

        # Copy the latest Valuation template
        cwd = pathlib.Path.cwd().resolve()
        try:
            template_folder_path = cwd / 'Stock_template'
            if pathlib.Path(template_folder_path).exists():
                path_list = [val_file_path for val_file_path in template_folder_path.iterdir()
                             if template_folder_path.is_dir() and val_file_path.is_file()]
                template_path_list = list(item for item in path_list if r.match(str(item)))
                if len(template_path_list) > 1 or len(template_path_list) == 0:
                    raise FileNotFoundError("The template file error", "temp_file")
            else:
                raise FileNotFoundError("The stock_template folder doesn't exist", "temp_folder")
        except FileNotFoundError as err:
            if err.args[1] == "temp_folder":
                print("The stock_template folder doesn't exist")
            if err.args[1] == "temp_file":
                print("The template file error")
        else:
            new_val_name = self.security_code + "_" + os.path.basename(template_path_list[0])
            new_val_path = cwd / new_val_name
            if not pathlib.Path(new_val_path).exists():
                print(f'Copying template to create {new_val_name}...')
                shutil.copy(template_path_list[0], new_val_path)
                new_bool = True
            # load and update the new valuation xlsx
            if os.path.exists(new_val_path):
                print(f'Updating {new_val_name}...')
                with xlwings.App(visible=False) as app:
                    xl_book = xlwings.Book(new_val_path)
                    self.update_dashboard(xl_book.sheets('Dashboard'), new_bool)
                    self.update_data(xl_book.sheets('Data'))
                    xl_book.save(new_val_name)
                    xl_book.close()
            else:
                raise FileNotFoundError("The valuation file error", "val_file")

    def update_dashboard(self, dash_sheet, new_bool=False):
        """Update the Dashboard sheet"""

        if new_bool:
            dash_sheet.range('C4').value = self.name
            dash_sheet.range('C5').value = datetime.today().strftime('%Y-%m-%d')
        dash_sheet.range('C3').value = self.security_code
        dash_sheet.range('H3').value = self.exchange
        dash_sheet.range('H12').value = self.report_currency
        dash_sheet.range('C6').value = self.next_earnings
        if pd.to_datetime(dash_sheet.range('C5').value) > pd.to_datetime(dash_sheet.range('C6').value):
            self.val_status = "Outdated"
        else:
            self.val_status = ""
        dash_sheet.range('E6').value = self.val_status
        dash_sheet.range('H4').value = self.price[0]
        dash_sheet.range('I4').value = self.price[1]
        dash_sheet.range('H5').value = self.shares
        dash_sheet.range('H13').value = scrap_mod.get_forex_rate(self.price[1], self.report_currency)

    def update_data(self, data_sheet):
        """Update the Data sheet"""

        data_sheet.range('C3').value = self.is_df.columns[0]  # last financial year
        # figures in
        figures_in = int((len(str(self.is_df.iloc[0, 0])) - 9) / 3 + 0.99) * 1000
        data_sheet.range('C4').value = figures_in
        # load income statement
        for i in range(len(self.is_df.columns)):
            data_sheet.range((7, i + 3)).value = int(self.is_df.iloc[0, i] / figures_in)
            data_sheet.range((9, i + 3)).value = int(self.is_df.iloc[1, i] / figures_in)
            data_sheet.range((11, i + 3)).value = int(self.is_df.iloc[2, i] / figures_in)
            data_sheet.range((17, i + 3)).value = int(self.is_df.iloc[3, i] / figures_in)
            data_sheet.range((18, i + 3)).value = int(self.is_df.iloc[4, i] / figures_in)
        # load balance sheet
        for i in range(1, len(self.bs_df.columns)):
            data_sheet.range((20, i + 3)).value = int(self.bs_df.iloc[0, i] / figures_in)
            data_sheet.range((21, i + 3)).value = int(self.bs_df.iloc[1, i] / figures_in)
            data_sheet.range((22, i + 3)).value = int(self.bs_df.iloc[2, i] / figures_in)
            data_sheet.range((23, i + 3)).value = int(self.bs_df.iloc[3, i] / figures_in)
            data_sheet.range((25, i + 3)).value = int(self.bs_df.iloc[4, i] / figures_in)
            data_sheet.range((26, i + 3)).value = int(self.bs_df.iloc[5, i] / figures_in)
            data_sheet.range((27, i + 3)).value = int(self.bs_df.iloc[6, i] / figures_in)
            data_sheet.range((28, i + 3)).value = int(self.bs_df.iloc[7, i] / figures_in)

    def export_statements(self):
        """Export the income statement and balance sheet"""

        self.is_df.to_csv(f'{self.security_code}_income_statement.csv', sep=',', encoding='utf-8')
        self.bs_df.to_csv(f'{self.security_code}_balance_sheet.csv', sep=',', encoding='utf-8')
