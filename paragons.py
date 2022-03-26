import start

def paragons():
    date = input("Enter a date of tank: ")
    start.wb.active['B14'] = date
    fuel = input("Enter amount of fuel: ")
    start.wb.active['C14'] = int(fuel)
    cost = input("Enter a fuel cost: ")
    start.wb.active['E14'] = int(cost)
    date1 = input("Enter a date of tank: ")
    start.wb.active['B155'] = date1
    fuel1 = input("Enter amount of fuel: ")
    start.wb.active['C15'] = int(fuel1)
    cost1 = input("Enter a fuel cost: ")
    start.wb.active['E15'] = int(cost1)
#TODO:czy trzeba robic po 10 zmiennych do kazdego miesiaca?
#TODO:car sort to ma byc tylko sortowanie a nie wybor
#TODO:funkcja w start , zobaczyc czy bedzie dzialac program na asmym `wb`
