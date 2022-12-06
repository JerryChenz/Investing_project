import openpyxl
import shutil
import pathlib
import os
import security_mod


def initiate_asset(dash_sheet):
    """initiate an asset from the valuation file"""

    a = security_mod.Asset(dash_sheet.cell(row=3, column=3).value)
    a.name = dash_sheet.cell(row=4, column=3).value
    a.exchange = dash_sheet.cell(row=3, column=8).value
    a.price = dash_sheet.cell(row=4, column=8).value
    a.excess_return = dash_sheet.cell(row=23, column=8).value

    return a


class Pipeline:
    """A pipeline with many assets"""

    def __init__(self):
        self.assets = []

    def load_opportunities(self):
        """Load the asset information from the opportunities folder"""

        # Copy the latest Valuation template
        opportunities_folder_path = pathlib.Path.cwd().resolve() / 'Opportunities'
        print(opportunities_folder_path)

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
            # wb = openpyxl.load_workbook(new_val_name)
            for p in opportunities_path_list:
                # load and update the new valuation xlsx
                dash_sheet = openpyxl.load_workbook(p)['Dashboard']
                self.assets.append(initiate_asset(dash_sheet))
