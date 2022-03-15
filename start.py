def asd():

    from openpyxl import Workbook, load_workbook
    import os

    os.chdir("/home/btn/PycharmProjects/pojazdy/2022")
    print(os.getcwd())

    wb = load_workbook('1.FORD COURIER.xlsx')
    ws = wb['styczen']


'''
def load():
    #wb = load_workbook('CITROEN JUMPER_MICHAL.xlsx')
    ws = wb.active

choose = input("Choose month from 1-12: ")
if choose == '1':
    #load()
    print(wb.sheetnames)
'''
