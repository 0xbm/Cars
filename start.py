from openpyxl import Workbook, load_workbook
import os

os.chdir("/home/btn/PycharmProjects/pojazdy/2022")
#print(os.getcwd())

wb = load_workbook('1.FORD COURIER.xlsx')
#ws = wb['styczen']