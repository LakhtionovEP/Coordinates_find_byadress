from geopy.geocoders import ArcGIS
import sys
import time
import os
import openpyxl
import re
wb = openpyxl.load_workbook('geo_input.xlsx')
sheet = wb.active

#Использование регулярной функции для выделения почтового индекса при его наличии

zipregex = re.compile(r'[A-Z0-9]+(\s|-)?(\d+$|\w+$)')
location = []
i = 2
while sheet['A' + str(i)].value != None:
    location.append(sheet['A' + str(i)].value)
    i += 1
geolocator = ArcGIS()
s, u, c, adr, zipc = 0, 0, 0, 0, 0
location_st2 = []
sheet['B1'] = 'Latitude'
sheet['C1'] = 'Longitude'
sheet['D1'] = 'CheckAddress1'
sheet['E1'] = 'CheckAddress2'
sheet['F1'] = 'ZipCode'

#Проверка, что результирующий файл закрыт

try:
    wb.save('geo_result.xlsx')
except PermissionError:
    print('Please close the resulting file' + ' geo_result.xlsx ' + 'before start')
    os.system('pause')
t = time.time()
for elem in range(len(location)):
    c += 1
    sys.stdout.write('\r%s' % c + ' from ' + str(len(location)) + ' addresses has been obtained ')
    sys.stdout.write(str(int((time.time() - t) // 60)) + ' minutes ' +
                     str(float('{:.1f}'.format((time.time() - t) % 60))) + ' seconds has been spent')
    sys.stdout.flush()
    global geol, x, y, adr1, adr2
    try:
        try:
            geol = geolocator.geocode(location[elem])
            x = float('{:.6f}'.format(geol.latitude))
            y = float('{:.6f}'.format(geol.longitude))
            s += 1
        except:
            try:
                time.sleep(1)
                geol = geolocator.geocode(location[elem])
                x = float('{:.6f}'.format(geol.latitude))
                y = float('{:.6f}'.format(geol.longitude))
                s += 1
            except:
                x = 'Can\'t find via service'
                y = 'Can\'t find via service'
                u += 1

#Обратное геокодирование для сравнения результатов

        try:
            adr1 = geol.address
            adr2 = str(geolocator.reverse(f'{x} {y}'))
            adr += 1
        except:
            try:
                time.sleep(1)
                adr1 = geol.address
                adr2 = str(geolocator.reverse(f'{x} {y}'))
                adr += 1
            except:
                try:
                    time.sleep(2)
                    adr1 = geol.address
                    adr2 = str(geolocator.reverse(f'{x} {y}'))
                    adr += 1
                except:
                    adr1 = 'Can\'t find via service'
                    adr2 = 'Can\'t find via service'
    finally:

        #Поиск почтового индекса

        try:
            zipr = str(geolocator.reverse(f'{geol.latitude} {geol.longitude}')).strip('\'').split(',')
            z = (zipregex.search(zipr[-2]).group())
            zipc += 1
        except:
            z = 'Can\'t find via service'
    sheet['B' + str(elem + 2)] = x
    sheet['C' + str(elem + 2)] = y
    sheet['D' + str(elem + 2)] = adr1
    sheet['E' + str(elem + 2)] = adr2
    sheet['F' + str(elem + 2)] = z
print()

#Статистика работы программы

print('Geocoding has finished')
print('Received Coordinates: ' + str(s))
print('Received Reverse Geocoding: ' + str(adr))
print('Zip-codes Recognised: ' + str(zipc))
print('Failed addresses: ' + str(u))
print('Overall time: ' + (str(int((time.time() - t) // 60)) + ' minutes ' +
                          str(float('{:.1f}'.format((time.time() - t) % 60)))) + ' seconds')
wb.save('geo_result.xlsx')
os.system('pause')
