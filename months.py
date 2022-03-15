def asd():
    import start

    mth = start.ws['C5']
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb['styczen']
        print(f"Wybrales {ws1}")
        x = input(f"{mth.value} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{mth.value} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{mth.value} fuel start: ")
        ws1['E10'] = int(z)

    start.wb.save("test.xlsx")


'''
elif choose == '2':
    mth = start.ws['C5']
    ws2 = start.wb['luty']
    x = input(f"{ws2.mth.value} odometer end: ")
    ws2['E8'] = int(x)
elif choose == '3':
    ws3 = start.wb['marzec']
    x = input(f"{mth.value} odometer end: ")
    ws3['E8'] = int(x)
elif choose == '4':
    ws4 = start.wb['kwiecien']
    x = input(f"{mth.value} odometer end: ")
    ws4['E8'] = int(x)
elif choose == '5':
    ws5 = start.wb['maj']
    x = input(f"{mth.value} odometer end: ")
    ws5['E8'] = int(x)
elif choose == '6':
    ws6 = start.wb['czerwiec']
    x = input(f"{mth.value} odometer end: ")
    ws6['E8'] = int(x)
elif choose == '7':
    ws7 = start.wb['lipiec']
    x = input(f"{mth.value} odometer end: ")
    ws7['E8'] = int(x)
elif choose == '8':
    ws8 = start.wb['sierpien']
    x = input(f"{mth.value} odometer end: ")
    ws8['E8'] = int(x)
elif choose == '9':
    ws9 = start.wb['wrzesien']
    x = input(f"{mth.value} odometer end: ")
    ws9['E8'] = int(x)
elif choose == '10':
    ws10 = start.wb['pazdziernik']
    x = input(f"{mth.value} odometer end: ")
    ws10['E8'] = int(x)
elif choose == '11':
    ws11 = start.wb['listopad']
    x = input(f"{mth.value} odometer end: ")
    ws11['E8'] = int(x)
elif choose == '12':
    ws12 = start.wb['grudzien']
    x = input(f"{mth.value} odometer end: ")
    ws12['E8'] = int(x)
'''
#start.wb.save("test.xlsx")
#import dat


'''
#luty
x = input(start.wb.get_sheet_by_name[1],"Luty odometer end: ")
ws2['E7'] = int(x)

#marzec

#kwiecien

#maj

#czerwiec

#lipiec

#sierpien

#wrzesien

#pazdziernik

#listopad

#grudzien
'''
