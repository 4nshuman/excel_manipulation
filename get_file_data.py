from os.path import join, dirname, abspath
import xlrd

def get_sheet_readable(filename, sheet_index, current_dir):
    fname = join(dirname(dirname(abspath(__file__))), current_dir, filename)
    # Open the workbook
    workbook = xlrd.open_workbook(fname)

    # List sheet names, and pull a sheet by name
    #sheet_names_list = xl_workbook.sheet_names()
    sheet = workbook.sheet_by_index(sheet_index)
    return sheet
