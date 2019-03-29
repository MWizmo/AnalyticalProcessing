import xlwt
import mysql.connector
import datetime


def bold_style():
    font = xlwt.Font()
    font.name = 'Arial'
    font.bold = True
    style = xlwt.XFStyle()
    style.font = font
    return style


def report_weight_city():
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    cursor.execute('SELECT * FROM ReportWeightByCity')
    rows = cursor.fetchall()
    file = xlwt.Workbook()
    style = bold_style()

    sheet = file.add_sheet('Report 1')
    sheet.write(0, 0, 'Year', style)
    sheet.write(0, 1, 'Month', style)
    sheet.write(0, 2, 'Day', style)
    sheet.write(0, 3, 'City', style)
    sheet.write(0, 4, 'Supplier', style)
    sheet.write(0, 5, 'Total weight', style)
    for i in range(0, len(rows)):
        flag = True
        for j in range(0, 6):
            if flag and rows[i][j] is None:
                sheet.write(i + 1, j, 'ИТОГО', style)
                flag = False
            else:
                sheet.write(i + 1, j, rows[i][j])
    now = datetime.datetime.today()
    date = now.strftime("%Y_%m_%d-%H_%M_%S")
    path = 'Reports\\WeightByCity\\Report ' + date + '.xls'
    file.save(path)
    return path


def report_price_city():
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    cursor.execute('SELECT * FROM ReportPriceByCity')
    rows = cursor.fetchall()
    file = xlwt.Workbook()
    style = bold_style()

    sheet = file.add_sheet('Report 1')
    sheet.write(0, 0, 'Year', style)
    sheet.write(0, 1, 'Month', style)
    sheet.write(0, 2, 'Day', style)
    sheet.write(0, 3, 'City', style)
    sheet.write(0, 4, 'Supplier', style)
    sheet.write(0, 5, 'Total price', style)
    for i in range(0, len(rows)):
        flag = True
        for j in range(0, 6):
            if flag and rows[i][j] is None:
                sheet.write(i + 1, j, 'ИТОГО', style)
                flag = False
            else:
                sheet.write(i + 1, j, rows[i][j])
    now = datetime.datetime.today()
    date = now.strftime("%Y_%m_%d-%H_%M_%S")
    path = 'Reports\\PriceByCity\\Report ' + date + '.xls'
    file.save(path)
    return path


def report_weight_price():
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    cursor.execute('SELECT * FROM ReportWeightByPrice')
    rows = cursor.fetchall()
    file = xlwt.Workbook()
    style = bold_style()

    sheet = file.add_sheet('Report 1')
    sheet.write(0, 0, 'Year', style)
    sheet.write(0, 1, 'Month', style)
    sheet.write(0, 2, 'Day', style)
    sheet.write(0, 3, 'Category', style)
    sheet.write(0, 4, 'Total weight', style)
    for i in range(0, len(rows)):
        flag = True
        for j in range(0, 5):
            if flag and rows[i][j] is None:
                sheet.write(i + 1, j, 'ИТОГО', style)
                flag = False
            else:
                sheet.write(i + 1, j, rows[i][j])
    now = datetime.datetime.today()
    date = now.strftime("%Y_%m_%d-%H_%M_%S")
    path = 'Reports\\WeightByPrice\\Report ' + date + '.xls'
    file.save(path)
    return path


def report_price_weight():
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    cursor.execute('SELECT * FROM ReportPriceByWeight')
    rows = cursor.fetchall()
    file = xlwt.Workbook()
    style = bold_style()

    sheet = file.add_sheet('Report 1')
    sheet.write(0, 0, 'Year', style)
    sheet.write(0, 1, 'Month', style)
    sheet.write(0, 2, 'Day', style)
    sheet.write(0, 3, 'Category', style)
    sheet.write(0, 4, 'Total price', style)
    for i in range(0, len(rows)):
        flag = True
        for j in range(0, 5):
            if flag and rows[i][j] is None:
                sheet.write(i + 1, j, 'ИТОГО', style)
                flag = False
            else:
                sheet.write(i + 1, j, rows[i][j])
    now = datetime.datetime.today()
    date = now.strftime("%Y_%m_%d-%H_%M_%S")
    path = 'Reports\\PriceByWeight\\Report ' + date + '.xls'
    file.save(path)
    return path
