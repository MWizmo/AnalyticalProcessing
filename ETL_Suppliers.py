import mysql.connector
import logging
import errors


def fillSuppliersTable(path, branch):
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    logging.basicConfig(filename="log.txt", level=logging.INFO)
    log = logging.getLogger('SuppliersLogger')
    log.info('Started\n\n')
    if branch == 1:
        import xlrd
        doc = xlrd.open_workbook(path)
        sheet = doc.sheet_by_index(0)
        succsessful_rows = 0
        for row in range(1, sheet.nrows):
            try:
                transformAndLoadSupplier(sheet.row(row)[0].value.replace('"', "'"), sheet.row(row)[1].value, sheet.row(row)[2].value,
                                     int(sheet.row(row)[3].value),branch, cursor, db)
                succsessful_rows += 1
            except errors.EmptyName:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустое название поставщика')
            except errors.EmptyCity:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Пустое название города, невозможно заменить наиболее частым значением для данного филиала')
            except errors.EmptyAddress:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустой адрес поставщика')
            except errors.BadValue:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Некорректный уровень риска, невозможно заменить средним значением для данного филиала')
            except mysql.connector.errors.IntegrityError:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Нарушение Unique')
        return 'Данные о поставщиках успешно добавлены в базу. Добавлено ' + str(succsessful_rows) + ' из ' + \
           str(sheet.nrows - 1) + ' записей. Подробности в файле log.txt'
    else:
        import pypyodbc
        connection = pypyodbc.win_connect_mdb(path)
        connection.cursor().execute('select * from Parts')
        a = connection.cursor().fetchall()
        connection.close()


def transformAndLoadSupplier(title, city, address, risk, branch, cursor, db):
    if title == '':
        raise errors.EmptyName
    if city == '':
        city = commonSupplierCity(title, branch, cursor)
        if city == '-':
            raise errors.EmptyCity
    if address == '':
        raise errors.EmptyAddress
    if risk < 1 or risk > 3:
        risk = averageRisk(city, branch, cursor)
        if risk is None:
            raise errors.BadValue
    cursor.execute('INSERT INTO Suppliers(Name,City,Address,Risk,Branch) VALUES ("'
                   + title + '","' + city + '","' + address + '",' + str(int(risk)) + ','+str(branch)+')')
    db.commit()


def commonSupplierCity(supplier, branch, cursor):
    cursor.execute('select City,COUNT(*) from Suppliers where Name="'+supplier+'" AND Branch='+str(branch)+' GROUP BY City order by COUNT(*) desc')
    city = cursor.fetchone()
    if city is None:
        return '-'
    else:
        return city[0]


def averageRisk(city, branch, cursor):
    cursor.execute('SELECT AVG(Risk) FROM Suppliers WHERE City="'+city+'" AND Branch='+str(branch))
    risks = cursor.fetchall()
    if len(risks) == 0:
        return -1
    else:
        return int(risks[0][0])
