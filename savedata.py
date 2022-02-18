import xlrd
from xlwt import Workbook
from xlutils.copy import copy
import datetime as date


def save_data(no_of_time_hand_detected, no_of_time_hand_crossed):
    try:
        today = date.today()
        today = str(today)
        # loc = (r'C:\Users\rahul.tripathi\Desktop\result.xls')

        rb = xlrd.open_workbook('result.xls')
        sheet = rb.sheet_by_index(0)
        sheet.cell_value(0, 0)

        # print(sheet.nrows)
        q = sheet.cell_value(sheet.nrows - 1, 1)

        rb = xlrd.open_workbook('result.xls')
        # rb = xlrd.open_workbook(loc)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)

        if q == today:
            w = sheet.cell_value(sheet.nrows - 1, 2)
            e = sheet.cell_value(sheet.nrows - 1, 3)
            w_sheet.write(sheet.nrows - 1, 2, w + no_of_time_hand_detected)
            w_sheet.write(sheet.nrows - 1, 3, e + no_of_time_hand_crossed)
            wb.save('result.xls')
        else:
            w_sheet.write(sheet.nrows, 0, sheet.nrows)
            w_sheet.write(sheet.nrows, 1, today)
            w_sheet.write(sheet.nrows, 2, no_of_time_hand_detected)
            w_sheet.write(sheet.nrows, 3, no_of_time_hand_crossed)
            wb.save('result.xls')
    except FileNotFoundError:
        today = date.today()
        today = str(today)

        # Workbook is created
        wb = Workbook()

        # add_sheet is used to create sheet.
        sheet = wb.add_sheet('Sheet 1')

        sheet.write(0, 0, 'Sl.No')
        sheet.write(0, 1, 'Date')
        sheet.write(0, 2, 'Number of times hand detected')
        sheet.write(0, 3, 'Number of times hand crossed')
        m = 1
        sheet.write(1, 0, m)
        sheet.write(1, 1, today)
        sheet.write(1, 2, no_of_time_hand_detected)
        sheet.write(1, 3, no_of_time_hand_crossed)

        wb.save('result.xls')
