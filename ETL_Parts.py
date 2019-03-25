import mysql.connector
import logging
import errors


def fillPartsTable(path, branch):
    db = mysql.connector.connect(host='localhost', database='Production',
                                   user='root', password='')
    cursor = db.cursor(buffered=True)
    logging.basicConfig(filename="log.txt", level=logging.INFO)
    log = logging.getLogger('PartsLogger')
    log.info('Started\n\n')
    if branch == 1:
        import xlrd
        doc = xlrd.open_workbook(path)
        sheet = doc.sheet_by_index(0)
        succsessful_rows = 0
        for row in range(1, sheet.nrows):
            try:
                transformAndLoadPart(sheet.row(row)[0].value, sheet.row(row)[1].value, sheet.row(row)[2].value, sheet.row(row)[3].value, branch, cursor, db)
                succsessful_rows += 1
            except errors.EmptyName:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустое название детали')
            except errors.EmptyCity:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Пустое название города, невозможно заменить наиболее частым значением для данного филиала')
            except errors.EmptyColor:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустой цвет детали')
            except errors.BadValue:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Отрицательный вес детали, невозможно заменить средним значением для данного филиала')
            except mysql.connector.errors.IntegrityError:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Нарушение Unique')
        return 'Данные о деталях успешно добавлены в базу. Добавлено '+str(succsessful_rows)+' из ' +\
               str(sheet.nrows-1)+' записей. Подробности в файле log.txt'
    else:
        pass


def transformAndLoadPart(title, city, color, weight, branch, cursor, db):
    if title == '':
        raise errors.EmptyName
    if city == '':
        city = commonCity(title, branch, cursor)
        if city == '-':
            raise errors.EmptyCity
    if color == '':
        raise errors.EmptyColor
    if float(weight) < 0:
        weight = averageWeight(city, branch, cursor)
        if weight == -1 or weight is None:
            raise errors.BadValue
    cursor.execute(
            'INSERT INTO Parts(Name,Color,City,Weight,Branch) VALUES ("' + title + '","' + color + '","' + city + '",' +str(
                weight) + ',' + str(branch)+')')
    db.commit()


def commonCity(part, branch, cursor):
    cursor.execute('select City,COUNT(*) from parts where Name="'+part+'" AND Branch='+str(branch)+' GROUP BY City order by COUNT(*) desc')
    city = cursor.fetchone()
    if city is None:
        return '-'
    else:
        return city[0]


def averageWeight(city, branch, cursor):
    cursor.execute('SELECT AVG(Weight) FROM Parts WHERE City="'+city+'" AND Branch='+str(branch))
    parts = cursor.fetchall()
    if len(parts) == 0:
        return -1
    else:
        return parts[0][0]


import pypyodbc
a = [x for x in pypyodbc.drivers() if x.startswith('Microsoft Access Driver')]
print(a)
# conn = pypyodbc.win_connect_mdb("Poduction.accdb")
# cur = conn.cursor()
# cur.execute('select * from Parts')
# a = cur.fetchall()
# print(a)