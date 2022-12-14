from argparse import ArgumentParser, RawDescriptionHelpFormatter
from os import listdir, getcwd, mkdir
from tqdm import tqdm
import zipfile

class UnlockExcel:
    def __init__(self, args):
        self.args = args
        self.path = self.args.path
        self.output = self.args.output
        # self.replace = self.args.replace
        # self.unhide = self.args.unhide
        self.file_type = "xlsx" # have original file type so it's converted back to the same either xlsx or xls

    def convert_to_zip(self):
        """Converts the xls/xlsx file into a zip file"""

        pass

    def convert_to_excel(self):
        """Converts the file back into xls/xlsx"""
        
        pass

    def remove_protection(self):
        """Removes protect tags from sheets and workbook"""

        pass

    def unhide_sheets(self):
        """Removes hide tags from sheets and rows/columns"""

        pass

    def unlock_file(self):
        """Unlocks the excel file"""

        pass

parser = ArgumentParser(
    formatter_class=RawDescriptionHelpFormatter, 
    description="""Unlocks excel spreadsheets and workbooks by going into the XML to remove the lock tags"""
)

parser.add_argument("-p", "--path", nargs="?", const=getcwd(), type=str, help="The path of the excel to unlock including the file name eg. path/to/file/filename.xlsx (path is cwd by default)")
parser.add_argument("-o", "--output", nargs="?", const=getcwd(), type=str, help="Output directory of the unlocked excel file")
# parser.add_argument("-r", "--replace", action="store_true", help="replaces the excel file instead of making a copy with _unlocked appended")
# parser.add_argument("-u", "--unhide", action="store_true", help="unhides all sheets and rows/columns")
args = parser.parse_args()
