import start

#daty i ilosci paliwa
mth = start.ws['C5']
choose = input("Choose number of paragons from 1-15: ")


if choose == '1':
    ws1 = start.wb['styczen']
    x = input(f"{mth.value} date 1st: ")
    ws1['B14'] = int(x)
    y = input(f"{mth.value} ilosc paliwa : ")
    ws1['C14'] = int(y)
elif choose == '2':
    ws2 = start.wb['luty']
    x = input(f"{mth.value} date 2nd: ")
    ws2['B14'] = int(x)
    y = input(f"{mth.value} ilosc paliwa : ")
    ws2['C14'] = int(y)


start.wb.save("test.xlsx")
'''Na chwile obecna WYBIERASZ liczbe paragonow do miesiaca STYCZEN 
ZROBIC FUNKCJE DLA KAZDEGO MIESIACA ?? 
'''