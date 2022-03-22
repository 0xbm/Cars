def ford_completing():
    import start
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb['styczen']
        sheet = start.wb.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        ws2 = start.wb['luty']
        sheet = start.wb.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        start.wb.save("test.xlsx")
    elif choose == '3':
        ws3 = start.wb['marzec']
        sheet = start.wb.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb['kwiecien']
        sheet = start.wb.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb['maj']
        sheet = start.wb.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb['czerwiec']
        sheet = start.wb.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb['lipiec']
        sheet = start.wb.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb['sierpien']
        sheet = start.wb.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb['wrzesien']
        sheet = start.wb.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb['pazdziernik']
        sheet = start.wb.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb['listopad']
        sheet = start.wb.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb['grudzien']
        sheet = start.wb.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb.save("1.FORD COURIER.xlsx")


def skoda_completing():
    import start
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb1['styczen']
        sheet = start.wb1.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        #mth = start.ws['C5']
        ws2 = start.wb1['luty']
        sheet = start.wb1.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb1['marzec']
        sheet = start.wb1.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb1['kwiecien']
        sheet = start.wb1.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb1['maj']
        sheet = start.wb1.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb1['czerwiec']
        sheet = start.wb1.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb1['lipiec']
        sheet = start.wb1.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb1['sierpien']
        sheet = start.wb1.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb1['wrzesien']
        sheet = start.wb1.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb1['pazdziernik']
        sheet = start.wb1.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb1['listopad']
        sheet = start.wb1.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb1['grudzien']
        sheet = start.wb1.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb1.save("2.SKODA FABIA.xlsx")


def skoda_2_completing():
    import start
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb2['styczen']
        sheet = start.wb2.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start.wb2['luty']
        sheet = start.wb2.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb2['marzec']
        sheet = start.wb2.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb2['kwiecien']
        sheet = start.wb2.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb2['maj']
        sheet = start.wb2.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb2['czerwiec']
        sheet = start.wb2.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb2['lipiec']
        sheet = start.wb2.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb2['sierpien']
        sheet = start.wb2.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb2['wrzesien']
        sheet = start.wb2.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb2['pazdziernik']
        sheet = start.wb2.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb2['listopad']
        sheet = start.wb2.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb2['grudzien']
        sheet = start.wb2.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb2.save("3.SKODA FABIA 2.xlsx")


def fiat_completing():
    import start
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb3['styczen']
        sheet = start.wb3.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start.wb3['luty']
        sheet = start.wb3.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb3['marzec']
        sheet = start.wb3.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb3['kwiecien']
        sheet = start.wb3.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb3['maj']
        sheet = start.wb3.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb3['czerwiec']
        sheet = start.wb3.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb3['lipiec']
        sheet = start.wb3.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb3['sierpien']
        sheet = start.wb3.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb3['wrzesien']
        sheet = start.wb3.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb3['pazdziernik']
        sheet = start.wb3.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb3['listopad']
        sheet = start.wb3.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb3['grudzien']
        sheet = start.wb3.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb3.save("3.FIAT PANDA.xlsx")





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
