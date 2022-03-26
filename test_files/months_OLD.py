import start_OLD


def ford_completing():
    choose = input("Choose month from 1-12: ")
    if choose == '1':
        ws1 = start_OLD.wb['styczen']
        sheet = start_OLD.wb.get_sheet_by_name('styczen')
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
    elif choose == '3':
        ws3 = start_OLD.wb['marzec']
        sheet = start.wb.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb['kwiecien']
        sheet = start.wb.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb['maj']
        sheet = start_OLD.wb.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start_OLD.wb['czerwiec']
        sheet = start_OLD.wb.get_sheet_by_name('czerwiec')
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
        ws9 = start_OLD.wb['wrzesien']
        sheet = start_OLD.wb.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb['pazdziernik']
        sheet = start.wb.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb['listopad']
        sheet = start.wb.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb['grudzien']
        sheet = start_OLD.wb.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb.save("1.FORD COURIER.xlsx")


def skoda_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start_OLD.wb1['styczen']
        sheet = start_OLD.wb1.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start.wb1['luty']
        sheet = start_OLD.wb1.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb1['marzec']
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
        sheet = start_OLD.wb1.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb1['czerwiec']
        sheet = start_OLD.wb1.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb1['lipiec']
        sheet = start_OLD.wb1.get_sheet_by_name('lipiec')
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
        sheet = start_OLD.wb1.get_sheet_by_name('pazdziernik')
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
        sheet = start_OLD.wb1.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb1.save("2.SKODA FABIA.xlsx")


def skoda_2_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb2['styczen']
        sheet = start_OLD.wb2.get_sheet_by_name('styczen')
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
        ws3 = start_OLD.wb2['marzec']
        sheet = start_OLD.wb2.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb2['kwiecien']
        sheet = start.wb2.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb2['maj']
        sheet = start_OLD.wb2.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb2['czerwiec']
        sheet = start_OLD.wb2.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb2['lipiec']
        sheet = start_OLD.wb2.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb2['sierpien']
        sheet = start_OLD.wb2.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb2['wrzesien']
        sheet = start_OLD.wb2.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb2['pazdziernik']
        sheet = start_OLD.wb2.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb2['listopad']
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
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb3['styczen']
        sheet = start_OLD.wb3.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start.wb3['luty']
        sheet = start_OLD.wb3.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb3['marzec']
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
        ws5 = start_OLD.wb3['maj']
        sheet = start_OLD.wb3.get_sheet_by_name('maj')
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
        ws7 = start_OLD.wb3['lipiec']
        sheet = start_OLD.wb3.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb3['sierpien']
        sheet = start_OLD.wb3.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb3['wrzesien']
        sheet = start.wb3.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb3['pazdziernik']
        sheet = start_OLD.wb3.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb3['listopad']
        sheet = start_OLD.wb3.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb3['grudzien']
        sheet = start_OLD.wb3.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb3.save("4.FIAT PANDA.xlsx")


def citroen_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start_OLD.wb4['styczen']
        sheet = start_OLD.wb4.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb4['luty']
        sheet = start.wb4.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb4['marzec']
        sheet = start_OLD.wb4.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb4['kwiecien']
        sheet = start.wb4.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb4['maj']
        sheet = start_OLD.wb4.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb4['czerwiec']
        sheet = start.wb4.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb4['lipiec']
        sheet = start_OLD.wb4.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb4['sierpien']
        sheet = start_OLD.wb4.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb4['wrzesien']
        sheet = start_OLD.wb4.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb4['pazdziernik']
        sheet = start_OLD.wb4.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb4['listopad']
        sheet = start_OLD.wb4.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb4['grudzien']
        sheet = start_OLD.wb4.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb4.save("5.CITROEN JUMPER.xlsx")


def daewoo_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start_OLD.wb5['styczen']
        sheet = start.wb5.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb5['luty']
        sheet = start.wb5.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb5['marzec']
        sheet = start.wb5.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb5['kwiecien']
        sheet = start_OLD.wb5.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb5['maj']
        sheet = start_OLD.wb5.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start_OLD.wb5['czerwiec']
        sheet = start.wb5.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb5['lipiec']
        sheet = start_OLD.wb5.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb5['sierpien']
        sheet = start_OLD.wb5.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb5['wrzesien']
        sheet = start.wb5.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb5['pazdziernik']
        sheet = start.wb5.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb5['listopad']
        sheet = start_OLD.wb5.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb5['grudzien']
        sheet = start.wb5.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb5.save("6.DAEWOO LUBLIN.xlsx")


def ford_2_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb6['styczen']
        sheet = start_OLD.wb6.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start.wb6['luty']
        sheet = start.wb6.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb6['marzec']
        sheet = start_OLD.wb6.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb6['kwiecien']
        sheet = start.wb6.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb6['maj']
        sheet = start.wb6.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb6['czerwiec']
        sheet = start_OLD.wb6.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb6['lipiec']
        sheet = start.wb6.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb6['sierpien']
        sheet = start_OLD.wb6.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb6['wrzesien']
        sheet = start_OLD.wb6.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb6['pazdziernik']
        sheet = start_OLD.wb6.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb6['listopad']
        sheet = start_OLD.wb6.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb6['grudzien']
        sheet = start_OLD.wb6.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb6.save("7.FORD TRANSIT.xlsx")


def ford_leasing_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb7['styczen']
        sheet = start_OLD.wb7.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb7['luty']
        sheet = start_OLD.wb7.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb7['marzec']
        sheet = start_OLD.wb7.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb7['kwiecien']
        sheet = start_OLD.wb7.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb7['maj']
        sheet = start_OLD.wb7.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb7['czerwiec']
        sheet = start_OLD.wb7.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb7['lipiec']
        sheet = start_OLD.wb7.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb7['sierpien']
        sheet = start.wb7.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb7['wrzesien']
        sheet = start_OLD.wb7.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb7['pazdziernik']
        sheet = start.wb7.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb7['listopad']
        sheet = start.wb7.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb7['grudzien']
        sheet = start.wb7.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb7.save("8.FORD TRANSIT LEASING.xlsx")


def farmtrac_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb8['styczen']
        sheet = start_OLD.wb8.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb8['luty']
        sheet = start.wb8.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb8['marzec']
        sheet = start.wb8.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb8['kwiecien']
        sheet = start_OLD.wb8.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb8['maj']
        sheet = start_OLD.wb8.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start_OLD.wb8['czerwiec']
        sheet = start.wb8.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb8['lipiec']
        sheet = start_OLD.wb8.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb8['sierpien']
        sheet = start_OLD.wb8.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start.wb8['wrzesien']
        sheet = start_OLD.wb8.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb8['pazdziernik']
        sheet = start_OLD.wb8.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb8['listopad']
        sheet = start_OLD.wb8.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb8['grudzien']
        sheet = start_OLD.wb8.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start.wb8.save("9.FARMTARAC.xlsx")


def ursus_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb9['styczen']
        sheet = start_OLD.wb9.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb9['luty']
        sheet = start_OLD.wb9.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb9['marzec']
        sheet = start_OLD.wb9.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb9['kwiecien']
        sheet = start_OLD.wb9.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb9['maj']
        sheet = start.wb9.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb9['czerwiec']
        sheet = start.wb9.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb9['lipiec']
        sheet = start_OLD.wb9.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start.wb9['sierpien']
        sheet = start_OLD.wb9.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb9['wrzesien']
        sheet = start_OLD.wb9.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb9['pazdziernik']
        sheet = start_OLD.wb9.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb9['listopad']
        sheet = start_OLD.wb9.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb9['grudzien']
        sheet = start.wb9.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb9.save("10.URSUS.xlsx")


def tym_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start_OLD.wb10['styczen']
        sheet = start_OLD.wb10.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb10['luty']
        sheet = start.wb10.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb10['marzec']
        sheet = start_OLD.wb10.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb10['kwiecien']
        sheet = start_OLD.wb10.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb10['maj']
        sheet = start_OLD.wb10.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb10['czerwiec']
        sheet = start.wb10.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb10['lipiec']
        sheet = start.wb10.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb10['sierpien']
        sheet = start_OLD.wb10.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb10['wrzesien']
        sheet = start_OLD.wb10.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb10['pazdziernik']
        sheet = start_OLD.wb10.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start.wb10['listopad']
        sheet = start_OLD.wb10.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb10['grudzien']
        sheet = start_OLD.wb10.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb10.save("11.TYM.xlsx")


def unimog_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start_OLD.wb11['styczen']
        sheet = start_OLD.wb11.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb11['luty']
        sheet = start_OLD.wb11.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb11['marzec']
        sheet = start_OLD.wb11.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb11['kwiecien']
        sheet = start.wb11.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb11['maj']
        sheet = start_OLD.wb11.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start.wb11['czerwiec']
        sheet = start_OLD.wb11.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb11['lipiec']
        sheet = start_OLD.wb11.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb11['sierpien']
        sheet = start_OLD.wb11.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb11['wrzesien']
        sheet = start.wb11.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb11['pazdziernik']
        sheet = start_OLD.wb11.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb11['listopad']
        sheet = start.wb11.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start_OLD.wb11['grudzien']
        sheet = start_OLD.wb11.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb11.save("12.UNIMOG.xlsx")


def unimog_2_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb12['styczen']
        sheet = start_OLD.wb12.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb12['luty']
        sheet = start.wb12.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start_OLD.wb12['marzec']
        sheet = start_OLD.wb12.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start_OLD.wb12['kwiecien']
        sheet = start.wb12.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start_OLD.wb12['maj']
        sheet = start_OLD.wb12.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start_OLD.wb12['czerwiec']
        sheet = start_OLD.wb12.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start.wb12['lipiec']
        sheet = start.wb12.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb12['sierpien']
        sheet = start.wb12.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb12['wrzesien']
        sheet = start_OLD.wb12.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start_OLD.wb12['pazdziernik']
        sheet = start_OLD.wb12.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb12['listopad']
        sheet = start_OLD.wb12.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb12['grudzien']
        sheet = start_OLD.wb12.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb12.save("13.UNIMOG_2.xlsx")


def noremat_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb13['styczen']
        sheet = start_OLD.wb13.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
    elif choose == '2':
        # mth = start.ws['C5']
        ws2 = start_OLD.wb13['luty']
        sheet = start_OLD.wb13.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
    elif choose == '3':
        ws3 = start.wb13['marzec']
        sheet = start.wb13.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
    elif choose == '4':
        ws4 = start.wb13['kwiecien']
        sheet = start.wb13.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
    elif choose == '5':
        ws5 = start.wb13['maj']
        sheet = start.wb13.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
    elif choose == '6':
        ws6 = start_OLD.wb13['czerwiec']
        sheet = start_OLD.wb13.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
    elif choose == '7':
        ws7 = start_OLD.wb13['lipiec']
        sheet = start_OLD.wb13.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
    elif choose == '8':
        ws8 = start_OLD.wb13['sierpien']
        sheet = start.wb13.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
    elif choose == '9':
        ws9 = start_OLD.wb13['wrzesien']
        sheet = start_OLD.wb13.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
    elif choose == '10':
        ws10 = start.wb13['pazdziernik']
        sheet = start_OLD.wb13.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
    elif choose == '11':
        ws11 = start_OLD.wb13['listopad']
        sheet = start_OLD.wb13.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
    else:
        ws12 = start.wb13['grudzien']
        sheet = start.wb13.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)

    start_OLD.wb13.save("14.NOREMAT.xlsx")
