import logging

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import datetime


def load_wb(filename):
    i = 0
    while True:
        if i > 5:
            raise RuntimeError("You tried too many times")
        try:
            f = open(filename, 'r+')
            break
        except IOError:
            logging.warning("User tried to open a file that is locked and is likely open in Excel.")
            input("Please close the open file in Excel and press enter. ")
            i += 1
    return load_workbook(filename)


def save_wb(filename, wb):
    wb.save(filename)


def initialize_workbook(wb: Workbook):
    ws = wb.create_sheet()
    data = ['Chuck Bland', 'ccbland@calpoly.edu', 12345, '(805) 756-7000', datetime.date(2023, 6, 5),
            datetime.date(2023, 6, 28), '=F2-TODAY()', 'Dale sterling lee dolan', 'dsdolan@calpoly.edu',
            'Chuck want to work on his Ham radio at home', 'Ham radio', 'F00BAR22']
    header = [i for i in dataframe().keys()]
    # header = ['Name', 'Email', 'EmplID', 'Phone', 'Checkout Start', 'Checkout End', 'Remaining', 'Advisor Name',
    #            'Advisor Email', 'Reason', 'Equipment', 'Equipment SN']
    ws.append(header)
    ws.append(data)
    tab = Table(displayName=next_table_name(wb), ref="A1:{}2".format(chr(ord('A') + len(data))))
    # for column, value in zip(tab.tableColumns, header):
    #     column.name = value
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True,
                           showColumnStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)
    return wb


def next_table_name(wb):
    tables = {}
    for sheet in wb.sheetnames:
        tables = tables | wb[sheet].tables
    base = 1
    while True:
        if base > 100:
            raise RuntimeError("Too many tables")
        if "Table{}".format(base) in tables:
            base += 1
        else:
            return "Table{}".format(base)


def dataframe():
    result = {"Name": None,
              'Email': None,
              'EmplID': None,
              'Phone': None,
              'Checkout Start': None,
              'Checkout End': None,
              'Remaining': None,
              'Advisor Name': None,
              'Advisor Email': None,
              'Reason': None,
              'Equipment': None,
              'Equipment SN': None
              }
    return result


def append_table(wb, data):
    ws = wb.active
    result = []
    for i in dataframe().keys():
        result.append(data[i])
    ws.append(result)
    return wb


def new_wb():
    wb = Workbook()
    return wb