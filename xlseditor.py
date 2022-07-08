import xlrd
import re
from datetime import date
from xlutils.copy import copy

book = xlrd.open_workbook('kpir.xls')
sheet = book.sheet_by_index(0)
date_input_regEx = '^d{4}\-(0[1-9]|1[012])\-(0[1-9]|[12][0-9]|3[01])$'
date_output_regEx =  '^([0-2][0-9]|(3)[0-1])-(((0)[0-9])|((1)[0-2]))-\d{4}$'
rb = xlrd.open_workbook("kpir_out.xls")
wb = copy(rb)
s = wb.get_sheet(0)

def convert_YYYYMMDD_to_DDMMYYYY():
    num_of_rows = sheet.nrows
    output = []
    # output - list of rows containing data standarized to DD/MM/YYYY
    for i in range(1,num_of_rows):
        inst_date = str(sheet.cell_value(rowx=i, colx = 0))
        inst_inf = str(sheet.cell_value(rowx=i, colx = 4))
        x = re.search(date_output_regEx,inst_date)
        if x:
            output.append([inst_date, inst_inf.split('.')[0]])
        else:
            right_date = ''.join([inst_date[8:10], '-', inst_date[5:7], '-', inst_date[0:4]])
            output.append([right_date, inst_inf])
    output.sort(key = lambda x: (x[0].split('-')[1], x[0].split('-')[0]))
    for i in range(len(output)):
        s.write(i,0, output[i][0])
        s.write(i,1, output[i][1])
        wb.save('kpir_out.xls')



convert_YYYYMMDD_to_DDMMYYYY()