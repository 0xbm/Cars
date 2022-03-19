def choose():
    import start
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb1['styczen']
        sheet = start.wb1.get_sheet_by_name('styczen')
        #print(f"Wybrales {mth.value}")
        x = input(f"{sheet.title} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title} fuel start: ")
        ws1['E10'] = int(z)
        start.wb1.save("test.xlsx")
    elif choose == '2':
        #mth = start.ws['C5']
        ws2 = start.wb['luty']
        sheet = start.wb.get_sheet_by_name('luty')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws2['E8'] = int(x)
        start.wb.save("test.xlsx")
    elif choose == '3':
        ws3 = start.wb['marzec']
        sheet = start.wb.get_sheet_by_name('marzec')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb['kwiecien']
        sheet = start.wb.get_sheet_by_name('kwiecien')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb['maj']
        sheet = start.wb.get_sheet_by_name('maj')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb['czerwiec']
        sheet = start.wb.get_sheet_by_name('czerwiec')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb['lipiec']
        sheet = start.wb.get_sheet_by_name('lipiec')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb['sierpien']
        sheet = start.wb.get_sheet_by_name('sierpien')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb['wrzesien']
        sheet = start.wb.get_sheet_by_name('wrzesien')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb['pazdziernik']
        sheet = start.wb.get_sheet_by_name('pazdziernik')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb['listopad']
        sheet = start.wb.get_sheet_by_name('listopad')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb['grudzien']
        sheet = start.wb.get_sheet_by_name('grudzien')
        print(f"Wybrales {sheet.title}")
        x = input(f"{sheet.title} odometer end: ")
        ws12['E8'] = int(x)

    #start.wb1.save("test.xlsx")




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
