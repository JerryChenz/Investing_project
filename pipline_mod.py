import openpyxl
import shutil
import pathlib
import os


class Pipeline:
    """A portfolio with many assets"""

    def __init__(self):
        self.assets = []
        self.files_path = ""
