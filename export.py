import xlwt
import datetime
import mysql.connector


def export_deliveries_by_dates():
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    cursor.execute('select ShipDate from Deliveries')
    dates = cursor.fetchall()
    file = xlwt.Workbook()
    sheet = file.add_sheet('Page 1')
    for i in range(0, len(dates)):
        string = ''
        cursor.execute('select p.Name from Deliveries d JOIN Parts p on p.ID=d.P_ID WHERE ShipDate="'+dates[i][0].strftime("%Y-%m-%d")+'"')
        parts = cursor.fetchall()
        for j in range(0, len(parts)):
            string += str(parts[j][0]) + ','
        string = string[:-1]
        sheet.write(i, 0, string)
    now = datetime.datetime.today()
    date = now.strftime("%Y_%m_%d-%H_%M_%S")
    path = 'Exports\\Export ' + date + '.xls'
    file.save(path)
    return path
