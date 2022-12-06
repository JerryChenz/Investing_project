import openpyxl
import xlwings
import pathlib
import security_mod


def initiate_asset(p):
    """initiate an asset from the valuation file"""

    # get the formula results using xlwings because openpyxl doesn't evaluate formula
    excel_app = xlwings.App(visible=False)
    excel_book = excel_app.books.open(p)
    excel_book.save()
    excel_book.close()
    excel_app.quit()
    wb = openpyxl.load_workbook(filename=p, data_only=True)

    dash_sheet = wb['Dashboard']
    print(dash_sheet.cell(row=21, column=8).value)

    a = security_mod.Asset(dash_sheet.cell(row=3, column=3).value)
    a.name = dash_sheet.cell(row=4, column=3).value
    a.exchange = dash_sheet.cell(row=3, column=8).value
    a.price = dash_sheet.cell(row=4, column=8).value
    a.ideal_price = dash_sheet.cell(row=21, column=3).value
    a.current_irr = dash_sheet.cell(row=21, column=8).value
    a.risk_premium = dash_sheet.cell(row=22, column=8).value

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

        monitor_sheet = wb['Monitor']

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
            r += 1

