import start

def paragons():
    date = input("Enter a date of tank: ")
    start.wb.active['B14'] = date
    fuel = input("Enter amount of fuel: ")
    start.wb.active['C14'] = int(fuel)
    cost = input("Enter a fuel cost: ")
    start.wb.active['E14'] = int(cost)
