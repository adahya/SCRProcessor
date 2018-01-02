from DBUpdate import FILENAME
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

if __name__ == '__main__':
    print("PHASE 1: Loading Excel Workbook %s" % FILENAME)
    try:
        wb=openpyxl.load_workbook(FILENAME)
        sheet = wb.get_sheet_by_name('TSSPM Report')
    except InvalidFileException as Error:
        print("Error : Trying to open Non-xlsx file %s" % FILENAME)