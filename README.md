usage: main.py [-h] -f FILE [-o OUTPUT] [-r] [-u]

Unlocks excel spreadsheets and workbooks by going into the XML to remove the lock tags.

optional arguments:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  The file path (eg. path/to/file/filename.xlsx).
  -o OUTPUT, --output OUTPUT
                        Output name and path of the unlocked excel file.
  -r, --replace         Replaces the excel file instead of making a copy with _unlocked appended
  -u, --unhide          unhides all sheets and rows/columns