import pandas
import collections
import openpyxl

#открываем файл с логами
excel_data = pandas.read_excel('logs.xlsx', sheet_name='log')

# открываем файл с отчетом и лист
rep = openpyxl.load_workbook(filename='report.xlsx')
sheet = rep['Лист1']

# выделяем столбец с браузерами в отдельную переменную
browsers = excel_data['Браузер']

# считаем кол-во повторений каждого браузера
collections_of_browsers = collections.Counter(browsers)

# выделяем 7 самых популярных браузеров
most_common_browsers = collections_of_browsers.most_common(7)

# выделяем столбец Купленные товары в отдельную переменную
products = excel_data['Купленные товары']

# Товары в строке, разделенны запятой
# например Кабель ZMI USB - microUSB (AL600) 1 м черный,Защитное стекло для Huawei Honor 30S,Мешок для обуви №1School синий K8547B
# создаем список all_products, в который добавляем неповторяющиеся товары, исключая "Ещё"
all_products = []
for product in products:
    string = product.split(',')
    for element in string:
        if 'Ещё' not in element:
            all_products.append(element)

# находим 7 самых популярных товаров
collections_of_products = collections.Counter(all_products)
most_common_products = collections_of_products.most_common(7)

#заполняем поля в отчете
for i in range(7):
#заполняем браузеры
    sheet['A'+str(5+i)] = most_common_browsers[i][0]
    sheet['B'+str(5+i)] = most_common_browsers[i][1]
#заполняем товары
    sheet['A'+str(19+i)] = most_common_products[i][0]
    sheet['B'+str(19+i)] = most_common_products[i][1]

# Вычисление количества посещений по месяцам

# Словарь с номерами месяцев и количеством посещений
month_dict = {'1': 0, '2': 0, '3': 0, '4': 0, '5': 0, '6': 0,'7': 0, '8': 0, '9': 0, '10': 0, '11': 0, '12': 0}

# Словарь с популярными браузерами и с посещениями по месяцам
browsers_month = {}
for i in range(7):
    browsers_month[most_common_browsers[i][0]] = month_dict.copy()

# Выбираем данные для обработки -  столбцы с датой посещения и браузером
data_plus_browsers = excel_data[['Дата посещения', 'Браузер']]

# Рассматриваем все строки, из каждой выделяем месяц и браузер,
# увеличивая кол-во посещений в соотвествующий месяц соответственного браузера в словаре browsers_month

# Создаем вспомогательный список с названиями популярных браузеров
excel_browsers = []
for brows in browsers_month.keys():
    excel_browsers.append(brows)

for i in range(data_plus_browsers.shape[0]):
    brows = data_plus_browsers.loc[i]['Браузер']
    month = str(data_plus_browsers.loc[i]['Дата посещения'].month)
    if brows in browsers_month.keys():
        browsers_month[brows][month] += 1

excel_rows = ['5', '6', '7', '8', '9', '10', '11']
excel_columns = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']

# Заполнение Количества посещений по месяцам
excel_browsers = []
for bros in browsers_month.keys():
    excel_browsers.append(bros)

for row in range(7):
    for col in range(12):
        sheet[excel_columns[col]+excel_rows[row]] = browsers_month[excel_browsers[row]][str(col+1)]

# Подсчитываем кол-во посещений по месяцам со всех браузеров
month_sum = month_dict.copy()
for v in browsers_month.values():
    for month, amount in v.items():
        month_sum[month] += amount

# Заполняем кол-во посещений по месяцам со всех браузеров
col = 0
for v in month_sum.values():
    sheet[excel_columns[col]+'12'] = v
    col += 1


# считаем продажи продуктов
data_plus_products = excel_data[['Дата посещения', 'Купленные товары']]

# Словарь с популярными товарами и с посещениями по месяцам
products_month = {}
for i in range(7):
    products_month[most_common_products[i][0]] = month_dict.copy()


for i in range(data_plus_products.shape[0]):
    month = str(data_plus_products.loc[i]['Дата посещения'].month)
    prod = data_plus_products.loc[i]['Купленные товары']
    string = prod.split(',')
    for element in string:
        if element in products_month.keys():
            products_month[element][month] += 1

# Заполнение товаров по месяцам
excel_rows = ['19', '20', '21', '22', '23', '24', '25']
excel_columns = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
excel_products = []
for prod in products_month.keys():
    excel_products.append(prod)

for row in range(7):
    for col in range(12):
        sheet[excel_columns[col]+excel_rows[row]] = products_month[excel_products[row]][str(col+1)]

month_sum = month_dict.copy()
for v in products_month.values():
    for month, amount in v.items():
        month_sum[month] += amount
col = 0
for v in month_sum.values():
    sheet[excel_columns[col]+'26'] = v
    col += 1

# Находим популярные и непопулярные товары у мужчин и женщин
sex_plus_products = excel_data[['Пол', 'Купленные товары']]

# Создадим и заполним список с товарами, купленными мужчинами и отдельный список с купленными женщинами
men_products = []
women_products = []

for i in range(sex_plus_products.shape[0]):
    sex = sex_plus_products.loc[i]['Пол']
    prod = sex_plus_products.loc[i]['Купленные товары']
    string = prod.split(',')
    for element in string:
        if 'Ещё' not in element:
            if sex == 'м':
                men_products.append(element)
            else:
                women_products.append(element)

# Находим самый популярный товар у мужчин и самый популярный у женщин
men_products_counter = collections.Counter(men_products)
women_products_counter = collections.Counter(women_products)

most_common_men_product = men_products_counter.most_common(1)[0][0]
most_common_women_product = women_products_counter.most_common(1)[0][0]

# Находим самый непопулярный товар у мужчин и самый популярный у женщин
len_men = len(men_products_counter)
len_women = len(women_products_counter)

result_men = men_products_counter.most_common()[:-(len_men+1):-1]
result_women = women_products_counter.most_common()[:-(len_women+1):-1]

less_common_men_product = result_men[0][0]
less_common_women_product = result_women[0][0]

# Заполняем соответствующие ячейки в таблице
sheet['B31'] = most_common_men_product
sheet['B32'] = most_common_women_product
sheet['B33'] = less_common_men_product
sheet['B34'] = less_common_women_product

# Сохраняем файл с отчетом
rep.save('report.xlsx')
