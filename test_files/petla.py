'''
for x in range(0, 5):
    print(input("Napisz cos/FOR"))
'''

#1
i = 0
while True:
    i += 1
    if i == 2:
        break
    print(input("Stan licznika na końcu miesiąca"))
    ws['F9'] = '12345'

#2 DZIAŁA
i = 0
while True:
    i += 1
    if i == 3: #liczba wykonywan kodu
        break
    x = input("Stan licznika na końcu miesiąca")
    ws['E9'] = int(x)
    y = input("Stan ")
    ws['E11'] = int(y)