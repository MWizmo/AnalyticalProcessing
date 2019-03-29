import mysql.connector
import logging
import errors


def fillDeliveriesTable(path, branch):
    db = mysql.connector.connect(host='localhost', database='Production',
                                 user='root', password='')
    cursor = db.cursor(buffered=True)
    logging.basicConfig(filename="log.txt", level=logging.INFO)
    log = logging.getLogger('DeliveriesLogger')
    log.info('Started\n\n')
    if branch == 2:
        import convertor
        zip_path = convertor.convert_accdb_to_xlsx(path)
        path = convertor.unzip_files(zip_path) + '\\Deliveries.xlsx'
    import xlrd
    doc = xlrd.open_workbook(path)
    sheet = doc.sheet_by_index(0)
    succsessful_rows = 0
    for row in range(1, sheet.nrows):
            date = sheet.row(row)[4].value
            if date == '':
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустое поле даты')
                continue
            date = xlrd.xldate.xldate_as_datetime(date, doc.datemode)
            date = str(date.year) + '-' + str(date.month) + '-' + str(date.day)
            try:
                transformAndLoadDelivery(sheet.row(row)[branch - 1 + 0].value, sheet.row(row)[branch - 1 + 1].value,
                                         sheet.row(row)[branch - 1 + 2].value, sheet.row(row)[branch - 1 + 3].value,
                                         date, branch, cursor, db)
                succsessful_rows += 1
            except errors.EmptyName:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Пустые ключи')
            except errors.BadQty:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Отрицательное количество деталей, невозможно заменить частым значением в рамках данного филиала')
            except errors.BadPrice:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Некорректное значение цены, невозможно заменить частым значением в рамках данного филиала')
            except mysql.connector.errors.IntegrityError:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) +
                          ': Нарушение внешнего ключа (ссылка на несуществующую запись)')
            except mysql.connector.errors.DatabaseError:
                log.error('Документ ' + path + '\nСтрока ' + str(row + 1) + ': Превышение допустимого веса поставки')
    return 'Данные о поставках успешно добавлены в базу. Добавлено ' + str(succsessful_rows) + ' из ' + \
           str(sheet.nrows - 1) + ' записей. Подробности в файле log.txt'


def transformAndLoadDelivery(s_id, p_id, qty, price, date, branch, cursor, db):
    if s_id == '' or p_id == '':
        raise errors.EmptyName
    qty = int(qty)
    if qty < 0:
        qty = commonQty(getCity(s_id, cursor), branch, cursor)
        if qty == '-':
            raise errors.BadQty
    if price == '':
        raise errors.BadPrice
    price = float(price)
    if price < 0:
        price = averagePrice(getCity(s_id, cursor), branch, cursor)
        if price == -1:
            raise errors.BadPrice
    cursor.execute('INSERT INTO Deliveries(S_ID,P_ID,Quantity,Price,ShipDate,Branch) VALUES ('+str(s_id)+','+str(p_id)+',' +
                   str(qty)+','+str(price)+',"'+date+'",'+str(branch)+')')
    db.commit()


def getCity(s_id, cursor):
    cursor.execute('SELECT s.City from Suppliers s join Deliveries d on s.ID=d.S_ID WHERE s.ID='+str(s_id))
    return cursor.fetchone()[0]


def commonQty(city, branch, cursor):
    cursor.execute('select s.City, d.Quantity from Deliveries d join Suppliers s on d.S_ID=s.ID where City="' + city + '" AND Branch=' + str(
        branch) + ' GROUP BY City order by d.Quantity desc')
    qty = cursor.fetchone()
    if qty is None:
        return '-'
    else:
        return qty[0]


def averagePrice(city, branch, cursor):
    cursor.execute('SELECT AVG(d.Price) FROM Deliveries d JOIN Suppliers s on d.S_ID=s.ID WHERE s.City="' + city + '" AND Branch=' + str(branch))
    prices = cursor.fetchall()
    if len(prices) == 0:
        return -1
    else:
        return float(prices[0][0])

