import unicodecsv
import xlwt
import os


INPUT_DIR = 'input'
OUTPUT_DIR = 'output'

def csv_to_xls(filename):
    outputName, fileExtension = os.path.splitext(filename)
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Statistics')

    f = open('%s/%s' % (INPUT_DIR, filename))

    reader = unicodecsv.reader(f, encoding='utf-8')

    for row_index, row in enumerate(reader):
        for column_index, cell in enumerate(row):
            ws.write(row_index, column_index, cell)
    wb.save('%s/%s.xls' % (OUTPUT_DIR, outputName))

    f.close()


for root,dirs,files in os.walk(INPUT_DIR):
    for file in files:
        if file.endswith(".csv") or file.endswith(".txt"):
            csv_to_xls(file)