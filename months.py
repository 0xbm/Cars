import start
import paragons

def ford_completing():
    choose = input("Choose month from 1-12: ")
    if choose == '1':
        start.wb.active = 0
        ws1 = start.wb['styczen']
        sheet = start.wb.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_ford()
    elif choose == '2':
        start.wb.active = 1
        ws2 = start.wb['luty']
        sheet = start.wb.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '3':
        start.wb.active = 2
        ws3 = start.wb['marzec']
        sheet = start.wb.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '4':
        start.wb.active = 3
        ws4 = start.wb['kwiecien']
        sheet = start.wb.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '5':
        start.wb.active = 4
        ws5 = start.wb['maj']
        sheet = start.wb.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '6':
        start.wb.active = 5
        ws6 = start.wb['czerwiec']
        sheet = start.wb.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '7':
        start.wb.active = 6
        ws7 = start.wb['lipiec']
        sheet = start.wb.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '8':
        start.wb.active = 7
        ws8 = start.wb['sierpien']
        sheet = start.wb.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '9':
        start.wb.active = 8
        ws9 = start.wb['wrzesien']
        sheet = start.wb.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '10':
        start.wb.active = 9
        ws10 = start.wb['pazdziernik']
        sheet = start.wb.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_ford()
    elif choose == '11':
        start.wb.active = 10
        ws11 = start.wb['listopad']
        sheet = start.wb.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_ford()
    else:
        start.wb.active = 11
        ws12 = start.wb['grudzien']
        sheet = start.wb.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_ford()

    start.wb.save("1.FORD COURIER.xlsx")


def skoda_completing():
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
        paragons.paragons_skoda()
    elif choose == '2':
        ws2 = start.wb1['luty']
        sheet = start.wb1.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '3':
        ws3 = start.wb1['marzec']
        sheet = start.wb1.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '4':
        ws4 = start.wb1['kwiecien']
        sheet = start.wb1.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '5':
        ws5 = start.wb1['maj']
        sheet = start.wb1.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '6':
        ws6 = start.wb1['czerwiec']
        sheet = start.wb1.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '7':
        ws7 = start.wb1['lipiec']
        sheet = start.wb1.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '8':
        ws8 = start.wb1['sierpien']
        sheet = start.wb1.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '9':
        ws9 = start.wb1['wrzesien']
        sheet = start.wb1.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '10':
        ws10 = start.wb1['pazdziernik']
        sheet = start.wb1.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_skoda()
    elif choose == '11':
        ws11 = start.wb1['listopad']
        sheet = start.wb1.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_skoda()
    else:
        ws12 = start.wb1['grudzien']
        sheet = start.wb1.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_skoda()

    start.wb1.save("2.SKODA FABIA.xlsx")


def skoda_2_completing():
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
        paragons.paragons_skoda_2()
    elif choose == '2':
        ws2 = start.wb2['luty']
        sheet = start.wb2.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '3':
        ws3 = start.wb2['marzec']
        sheet = start.wb2.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '4':
        ws4 = start.wb2['kwiecien']
        sheet = start.wb2.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '5':
        ws5 = start.wb2['maj']
        sheet = start.wb2.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '6':
        ws6 = start.wb2['czerwiec']
        sheet = start.wb2.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '7':
        ws7 = start.wb2['lipiec']
        sheet = start.wb2.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '8':
        ws8 = start.wb2['sierpien']
        sheet = start.wb2.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '9':
        ws9 = start.wb2['wrzesien']
        sheet = start.wb2.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '10':
        ws10 = start.wb2['pazdziernik']
        sheet = start.wb2.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_skoda_2()
    elif choose == '11':
        ws11 = start.wb2['listopad']
        sheet = start.wb2.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_skoda_2()
    else:
        ws12 = start.wb2['grudzien']
        sheet = start.wb2.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_skoda_2()

    start.wb2.save("3.SKODA FABIA 2.xlsx")


def fiat_completing():
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
        paragons.paragons_fiat()
    elif choose == '2':
        ws2 = start.wb3['luty']
        sheet = start.wb3.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '3':
        ws3 = start.wb3['marzec']
        sheet = start.wb3.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '4':
        ws4 = start.wb3['kwiecien']
        sheet = start.wb3.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '5':
        ws5 = start.wb3['maj']
        sheet = start.wb3.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '6':
        ws6 = start.wb3['czerwiec']
        sheet = start.wb3.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '7':
        ws7 = start.wb3['lipiec']
        sheet = start.wb3.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '8':
        ws8 = start.wb3['sierpien']
        sheet = start.wb3.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '9':
        ws9 = start.wb3['wrzesien']
        sheet = start.wb3.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '10':
        ws10 = start.wb3['pazdziernik']
        sheet = start.wb3.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_fiat()
    elif choose == '11':
        ws11 = start.wb3['listopad']
        sheet = start.wb3.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_fiat()
    else:
        ws12 = start.wb3['grudzien']
        sheet = start.wb3.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_fiat()

    start.wb3.save("4.FIAT PANDA.xlsx")


def citroen_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb4['styczen']
        sheet = start.wb4.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_citroen()
    elif choose == '2':
        ws2 = start.wb4['luty']
        sheet = start.wb4.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '3':
        ws3 = start.wb4['marzec']
        sheet = start.wb4.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '4':
        ws4 = start.wb4['kwiecien']
        sheet = start.wb4.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '5':
        ws5 = start.wb4['maj']
        sheet = start.wb4.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '6':
        ws6 = start.wb4['czerwiec']
        sheet = start.wb4.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '7':
        ws7 = start.wb4['lipiec']
        sheet = start.wb4.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '8':
        ws8 = start.wb4['sierpien']
        sheet = start.wb4.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '9':
        ws9 = start.wb4['wrzesien']
        sheet = start.wb4.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '10':
        ws10 = start.wb4['pazdziernik']
        sheet = start.wb4.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_citroen()
    elif choose == '11':
        ws11 = start.wb4['listopad']
        sheet = start.wb4.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_citroen()
    else:
        ws12 = start.wb4['grudzien']
        sheet = start.wb4.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_citroen()

    start.wb4.save("5.CITROEN JUMPER.xlsx")


def daewoo_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb5['styczen']
        sheet = start.wb5.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_daewoo()
    elif choose == '2':
        ws2 = start.wb5['luty']
        sheet = start.wb5.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '3':
        ws3 = start.wb5['marzec']
        sheet = start.wb5.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '4':
        ws4 = start.wb5['kwiecien']
        sheet = start.wb5.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '5':
        ws5 = start.wb5['maj']
        sheet = start.wb5.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '6':
        ws6 = start.wb5['czerwiec']
        sheet = start.wb5.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '7':
        ws7 = start.wb5['lipiec']
        sheet = start.wb5.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '8':
        ws8 = start.wb5['sierpien']
        sheet = start.wb5.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '9':
        ws9 = start.wb5['wrzesien']
        sheet = start.wb5.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '10':
        ws10 = start.wb5['pazdziernik']
        sheet = start.wb5.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_daewoo()
    elif choose == '11':
        ws11 = start.wb5['listopad']
        sheet = start.wb5.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_daewoo()
    else:
        ws12 = start.wb5['grudzien']
        sheet = start.wb5.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_daewoo()

    start.wb5.save("6.DAEWOO LUBLIN.xlsx")


def ford_2_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb6['styczen']
        sheet = start.wb6.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_ford_2()
    elif choose == '2':
        ws2 = start.wb6['luty']
        sheet = start.wb6.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '3':
        ws3 = start.wb6['marzec']
        sheet = start.wb6.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '4':
        ws4 = start.wb6['kwiecien']
        sheet = start.wb6.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '5':
        ws5 = start.wb6['maj']
        sheet = start.wb6.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '6':
        ws6 = start.wb6['czerwiec']
        sheet = start.wb6.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '7':
        ws7 = start.wb6['lipiec']
        sheet = start.wb6.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '8':
        ws8 = start.wb6['sierpien']
        sheet = start.wb6.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '9':
        ws9 = start.wb6['wrzesien']
        sheet = start.wb6.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '10':
        ws10 = start.wb6['pazdziernik']
        sheet = start.wb6.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_ford_2()
    elif choose == '11':
        ws11 = start.wb6['listopad']
        sheet = start.wb6.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_ford_2()
    else:
        ws12 = start.wb6['grudzien']
        sheet = start.wb6.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_ford_2()

    start.wb6.save("7.FORD TRANSIT.xlsx")


def ford_leasing_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb7['styczen']
        sheet = start.wb7.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_ford_leasing()
    elif choose == '2':
        ws2 = start.wb7['luty']
        sheet = start.wb7.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '3':
        ws3 = start.wb7['marzec']
        sheet = start.wb7.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '4':
        ws4 = start.wb7['kwiecien']
        sheet = start.wb7.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '5':
        ws5 = start.wb7['maj']
        sheet = start.wb7.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '6':
        ws6 = start.wb7['czerwiec']
        sheet = start.wb7.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '7':
        ws7 = start.wb7['lipiec']
        sheet = start.wb7.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '8':
        ws8 = start.wb7['sierpien']
        sheet = start.wb7.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '9':
        ws9 = start.wb7['wrzesien']
        sheet = start.wb7.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '10':
        ws10 = start.wb7['pazdziernik']
        sheet = start.wb7.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_ford_leasing()
    elif choose == '11':
        ws11 = start.wb7['listopad']
        sheet = start.wb7.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_ford_leasing()
    else:
        ws12 = start.wb7['grudzien']
        sheet = start.wb7.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_ford_leasing()

    start.wb7.save("8.FORD TRANSIT LEASING.xlsx")


def farmtrac_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb8['styczen']
        sheet = start.wb8.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_farmtrac()
    elif choose == '2':
        ws2 = start.wb8['luty']
        sheet = start.wb8.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '3':
        ws3 = start.wb8['marzec']
        sheet = start.wb8.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '4':
        ws4 = start.wb8['kwiecien']
        sheet = start.wb8.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '5':
        ws5 = start.wb8['maj']
        sheet = start.wb8.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '6':
        ws6 = start.wb8['czerwiec']
        sheet = start.wb8.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '7':
        ws7 = start.wb8['lipiec']
        sheet = start.wb8.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '8':
        ws8 = start.wb8['sierpien']
        sheet = start.wb8.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '9':
        ws9 = start.wb8['wrzesien']
        sheet = start.wb8.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '10':
        ws10 = start.wb8['pazdziernik']
        sheet = start.wb8.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_farmtrac()
    elif choose == '11':
        ws11 = start.wb8['listopad']
        sheet = start.wb8.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_farmtrac()
    else:
        ws12 = start.wb8['grudzien']
        sheet = start.wb8.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_farmtrac()

    start.wb8.save("9.FARMTARAC.xlsx")


def ursus_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb9['styczen']
        sheet = start.wb9.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_ursus()
    elif choose == '2':
        ws2 = start.wb9['luty']
        sheet = start.wb9.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '3':
        ws3 = start.wb9['marzec']
        sheet = start.wb9.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '4':
        ws4 = start.wb9['kwiecien']
        sheet = start.wb9.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '5':
        ws5 = start.wb9['maj']
        sheet = start.wb9.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '6':
        ws6 = start.wb9['czerwiec']
        sheet = start.wb9.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '7':
        ws7 = start.wb9['lipiec']
        sheet = start.wb9.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '8':
        ws8 = start.wb9['sierpien']
        sheet = start.wb9.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '9':
        ws9 = start.wb9['wrzesien']
        sheet = start.wb9.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '10':
        ws10 = start.wb9['pazdziernik']
        sheet = start.wb9.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_ursus()
    elif choose == '11':
        ws11 = start.wb9['listopad']
        sheet = start.wb9.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_ursus()
    else:
        ws12 = start.wb9['grudzien']
        sheet = start.wb9.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_ursus()

    start.wb9.save("10.URSUS.xlsx")


def tym_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb10['styczen']
        sheet = start.wb10.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_tym()
    elif choose == '2':
        ws2 = start.wb10['luty']
        sheet = start.wb10.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '3':
        ws3 = start.wb10['marzec']
        sheet = start.wb10.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '4':
        ws4 = start.wb10['kwiecien']
        sheet = start.wb10.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '5':
        ws5 = start.wb10['maj']
        sheet = start.wb10.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '6':
        ws6 = start.wb10['czerwiec']
        sheet = start.wb10.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '7':
        ws7 = start.wb10['lipiec']
        sheet = start.wb10.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '8':
        ws8 = start.wb10['sierpien']
        sheet = start.wb10.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '9':
        ws9 = start.wb10['wrzesien']
        sheet = start.wb10.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '10':
        ws10 = start.wb10['pazdziernik']
        sheet = start.wb10.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_tym()
    elif choose == '11':
        ws11 = start.wb10['listopad']
        sheet = start.wb10.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_tym()
    else:
        ws12 = start.wb10['grudzien']
        sheet = start.wb10.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_tym()

    start.wb10.save("11.TYM.xlsx")


def unimog_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb11['styczen']
        sheet = start.wb11.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_unimog()
    elif choose == '2':
        ws2 = start.wb11['luty']
        sheet = start.wb11.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '3':
        ws3 = start.wb11['marzec']
        sheet = start.wb11.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '4':
        ws4 = start.wb11['kwiecien']
        sheet = start.wb11.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '5':
        ws5 = start.wb11['maj']
        sheet = start.wb11.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '6':
        ws6 = start.wb11['czerwiec']
        sheet = start.wb11.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '7':
        ws7 = start.wb11['lipiec']
        sheet = start.wb11.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '8':
        ws8 = start.wb11['sierpien']
        sheet = start.wb11.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '9':
        ws9 = start.wb11['wrzesien']
        sheet = start.wb11.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '10':
        ws10 = start.wb11['pazdziernik']
        sheet = start.wb11.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_unimog()
    elif choose == '11':
        ws11 = start.wb11['listopad']
        sheet = start.wb11.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_unimog()
    else:
        ws12 = start.wb11['grudzien']
        sheet = start.wb11.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_unimog()

    start.wb11.save("12.UNIMOG.xlsx")


def unimog_2_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb12['styczen']
        sheet = start.wb12.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_unimog_2()
    elif choose == '2':
        ws2 = start.wb12['luty']
        sheet = start.wb12.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '3':
        ws3 = start.wb12['marzec']
        sheet = start.wb12.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '4':
        ws4 = start.wb12['kwiecien']
        sheet = start.wb12.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '5':
        ws5 = start.wb12['maj']
        sheet = start.wb12.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '6':
        ws6 = start.wb12['czerwiec']
        sheet = start.wb12.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '7':
        ws7 = start.wb12['lipiec']
        sheet = start.wb12.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '8':
        ws8 = start.wb12['sierpien']
        sheet = start.wb12.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '9':
        ws9 = start.wb12['wrzesien']
        sheet = start.wb12.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '10':
        ws10 = start.wb12['pazdziernik']
        sheet = start.wb12.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_unimog_2()
    elif choose == '11':
        ws11 = start.wb12['listopad']
        sheet = start.wb12.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_unimog_2()
    else:
        ws12 = start.wb12['grudzien']
        sheet = start.wb12.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_unimog_2()

    start.wb12.save("13.UNIMOG_2.xlsx")


def noremat_completing():
    choose = input("Choose month from 1-12: ")

    if choose == '1':
        ws1 = start.wb13['styczen']
        sheet = start.wb13.get_sheet_by_name('styczen')
        x = input(f"{sheet.title.capitalize()} odometer start: ")
        ws1['E7'] = int(x)
        y = input(f"{sheet.title.capitalize()} odometer end: ")
        ws1['E8'] = int(y)
        z = input(f"{sheet.title.capitalize()} fuel start: ")
        ws1['E10'] = int(z)
        paragons.paragons_noremat()
    elif choose == '2':
        ws2 = start.wb13['luty']
        sheet = start.wb13.get_sheet_by_name('luty')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws2['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '3':
        ws3 = start.wb13['marzec']
        sheet = start.wb13.get_sheet_by_name('marzec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws3['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '4':
        ws4 = start.wb13['kwiecien']
        sheet = start.wb13.get_sheet_by_name('kwiecien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws4['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '5':
        ws5 = start.wb13['maj']
        sheet = start.wb13.get_sheet_by_name('maj')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws5['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '6':
        ws6 = start.wb13['czerwiec']
        sheet = start.wb13.get_sheet_by_name('czerwiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws6['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '7':
        ws7 = start.wb13['lipiec']
        sheet = start.wb13.get_sheet_by_name('lipiec')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws7['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '8':
        ws8 = start.wb13['sierpien']
        sheet = start.wb13.get_sheet_by_name('sierpien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws8['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '9':
        ws9 = start.wb13['wrzesien']
        sheet = start.wb13.get_sheet_by_name('wrzesien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws9['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '10':
        ws10 = start.wb13['pazdziernik']
        sheet = start.wb13.get_sheet_by_name('pazdziernik')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws10['E8'] = int(x)
        paragons.paragons_noremat()
    elif choose == '11':
        ws11 = start.wb13['listopad']
        sheet = start.wb13.get_sheet_by_name('listopad')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws11['E8'] = int(x)
        paragons.paragons_noremat()
    else:
        ws12 = start.wb13['grudzien']
        sheet = start.wb13.get_sheet_by_name('grudzien')
        print(f"You choose {sheet.title}")
        x = input(f"{sheet.title.capitalize()} odometer end: ")
        ws12['E8'] = int(x)
        paragons.paragons_noremat()

    start.wb13.save("14.NOREMAT.xlsx")
