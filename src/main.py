from argparse import ArgumentParser, RawDescriptionHelpFormatter
from os import listdir, getcwd, remove
from tqdm import tqdm
import zipfile
import shutil
import re


class UnlockExcel:
    def __init__(self, args):
        self.args = args
        self.file = self.args.file
        self.output = self.args.output
        self.replace = self.args.replace
        self.unhide = self.args.unhide
        self.file_type = "xlsx" # have original file type so it's converted back to the same either xlsx or xls.
        self.cwd = getcwd()

    def convert_to_zip(self):
        """Converts the xls/xlsx file into a zip file and extracts."""

        self.file_type = self.file.split('.')[1]

        zip_name = f"{self.file.split('.')[0]}.zip"
        shutil.copy(self.file, zip_name) # create copy zip file
        
        with zipfile.ZipFile(zip_name, 'r') as zip_ref:
            zip_ref.extractall(f"{self.cwd}/temp_{zip_name.split('.')[0]}") # unzip file to edit
            self.temp_dir = f"{self.cwd}/temp_{zip_name.split('.')[0]}"
        
        remove(zip_name)
        if self.replace:
            remove(self.file)

    def convert_to_excel(self):
        """Converts back into xls/xlsx and deletes directory."""
        
        pass

    def remove_protection(self):
        """Removes protect tags from sheets and workbook."""

        # with open(f"{self.temp_dir}/xl/workbook.xml") as workbook_file:
        #     xml = workbook_file.read()

        sheets_dir = f"{self.temp_dir}/xl/worksheets"
        sheets = listdir(sheets_dir)
        for sheet in sheets:
            with open(f"{sheets_dir}/{sheet}", "r") as worksheet_file:
                xml = worksheet_file.read()

            regex = r"<\bsheetProtection\b.+?/>"
            xml = re.sub(regex, "", xml)

            with open(f"{sheets_dir}/{sheet}", "w") as worksheet_file:
                worksheet_file.write(xml)


    def unhide_sheets(self):
        """Removes hide tags from sheets and rows/columns."""

        # in workbook.xml the sheets are hidden by state="hidden" in tags
        pass

    def unlock_file(self):
        """Unlocks the excel file."""

        self.convert_to_zip()
        self.remove_protection()

        if self.unhide:
            self.unhide_sheets()
        
        self.convert_to_excel()

parser = ArgumentParser(
    formatter_class=RawDescriptionHelpFormatter, 
    description="""Unlocks excel spreadsheets and workbooks by going into the XML to remove the lock tags."""
)

parser.add_argument("-f", "--file", required=True, type=str, help="The file path (eg. path/to/file/filename.xlsx).")
parser.add_argument("-o", "--output", const=None, required=False, type=str, help="Output name and path of the unlocked excel file.")
parser.add_argument("-r", "--replace", action="store_true", help="Replaces the excel file instead of making a copy with _unlocked appended")
parser.add_argument("-u", "--unhide", action="store_true", help="unhides all sheets and rows/columns")
args = parser.parse_args()

instance = UnlockExcel(args)
instance.unlock_file()