import openpyxl
import xlwings
import pathlib
import security_mod
import scrap_mod
import yfinance
import pandas as pd


def update_stocks_val(dash_sheet):
    """Update the stock valuations in the pipeline folder"""

    ticker_info = yfinance.Ticker(dash_sheet.cell(row=3, column=3).value).info

    if pd.to_datetime(dash_sheet.cell(row=5, column=3).value) > pd.to_datetime(dash_sheet.cell(row=5, column=3).value):
        dash_sheet.cell(row=6, column=5).value = "Outdated"
    else:
        dash_sheet.cell(row=6, column=5).value = ""
    dash_sheet.cell(row=4, column=8).value = ticker_info['currentPrice']
    # dash_sheet.cell(row=5, column=8).value = ticker_info['sharesOutstanding']
    dash_sheet.cell(row=13, column=8).value = scrap_mod.get_forex_rate(dash_sheet.cell(row=4, column=9).value,
                                                                       dash_sheet.cell(row=12, column=8).value)


def initiate_asset(p):
    """initiate an asset from the valuation file"""

    # get the formula results using xlwings because openpyxl doesn't evaluate formula
    excel_app = xlwings.App(visible=False)
    excel_book = excel_app.books.open(p)
    excel_book.save()
    excel_book.close()
    excel_app.quit()
    wb = openpyxl.load_workbook(filename=p, data_only=True)

    # Update the stock_valuations files first
    dash_sheet = wb['Dashboard']
    update_stocks_val(dash_sheet)
    wb.save(p)

    a = security_mod.Asset(dash_sheet.cell(row=3, column=3).value)
    a.name = dash_sheet.cell(row=4, column=3).value
    a.exchange = dash_sheet.cell(row=3, column=8).value
    a.price = dash_sheet.cell(row=4, column=8).value
    a.ideal_price = dash_sheet.cell(row=21, column=3).value
    a.current_irr = dash_sheet.cell(row=21, column=8).value
    a.risk_premium = dash_sheet.cell(row=22, column=8).value
    a.val_status = dash_sheet.cell(row=6, column=5).value

    return a


class Pipeline:
    """A pipeline with many assets"""

    def __init__(self):
        self.assets = []

    def load_opportunities(self):
        """Load the asset information from the opportunities folder"""

        # Copy the latest Valuation template
        opportunities_folder_path = pathlib.Path.cwd().resolve() / 'Opportunities'

        try:
            if pathlib.Path(opportunities_folder_path).exists():
                opportunities_path_list = [val_file_path for val_file_path in opportunities_folder_path.iterdir()
                                           if opportunities_folder_path.is_dir() and val_file_path.is_file()]
                if len(opportunities_path_list) == 0:
                    raise FileNotFoundError("No opportunity file", "opp_file")
            else:
                raise FileNotFoundError("The opportunities folder doesn't exist", "opp_folder")
        except FileNotFoundError as err:
            if err.args[1] == "opp_folder":
                print("The opportunities folder doesn't exist")
            if err.args[1] == "opp_file":
                print("No opportunity file", "opp_file")
        else:
            # load and update the new valuation xlsx
            for p in opportunities_path_list:
                # load and update the new valuation xlsx
                self.assets.append(initiate_asset(p))
        # load the opportunities
        monitor_file_path = opportunities_folder_path / 'Pipeline_monitor' / 'Pipeline_monitor.xlsx'
        monitor_wb = openpyxl.load_workbook(monitor_file_path)
        self.update_monitor(monitor_wb)
        monitor_wb.save(monitor_file_path)

    def update_monitor(self, wb):
        """update the Pipeline_monitor file"""

        # Clear existing data
        monitor_sheet = wb['Monitor']
        for c in monitor_sheet['B5':'I200']:
            for cell in c:
                cell.value = None

        r = 5
        for a in self.assets:
            monitor_sheet.cell(row=r, column=2).value = a.security_code
            monitor_sheet.cell(row=r, column=3).value = a.name
            monitor_sheet.cell(row=r, column=4).value = a.exchange
            monitor_sheet.cell(row=r, column=5).value = a.price
            monitor_sheet.cell(row=r, column=6).value = a.current_irr
            monitor_sheet.cell(row=r, column=7).value = a.risk_premium
            monitor_sheet.cell(row=r, column=8).value = f'=F{r}-G{r}'
            monitor_sheet.cell(row=r, column=9).value = a.ideal_price
            monitor_sheet.cell(row=r, column=10).value = a.val_status
            r += 1
