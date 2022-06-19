from natsort import natsorted
import months
import glob
import os
import re


class DisplayMenu:
    def sort_years(self):
        path = os.getcwd()
        folders_sort = ([name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name))])
        folders_sort.reverse()
        folders_exception = re.compile(r".*(venv|.idea)")
        print([folders_exception.sub("", x) for x in folders_sort])

    def choose_year(self):
        choose = input("Choose year: 1.2022, 2.2023: ")
        if choose == "1":
            os.chdir("/Users/btn/PycharmProjects/cars/2022")
            print("You choosed year: 2022")
            t.choose_cars()
        if choose == "2":
            os.chdir("/Users/btn/PycharmProjects/cars/2023")
            print("You choosed year: 2023")
            t.choose_cars()

    def choose_cars(self):
        path = (os.getcwd())
        files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
        folders_show = re.compile(r".*(/2022/|/2023/)")
        files_sort = natsorted([folders_show.sub('', x).strip() for x in files])
        print(files_sort)
        choose = input("Select vehicle number: ")
        if choose == '1':
            print("You selected: ", files_sort[0])
            months.ford_completing()
        if choose == '2':
            print("You selected: ", files_sort[1])
            months.skoda_completing()
        if choose == '3':
            print("You selected: ", files_sort[2])
            months.skoda_2_completing()
        if choose == '4':
            print("You selected: ", files_sort[3])
            months.fiat_completing()
        if choose == '5':
            print("You selected: ", files_sort[4])
            months.citroen_completing()
        if choose == '6':
            print("You selected: ", files_sort[5])
            months.daewoo_completing()
        if choose == '7':
            print("You selected: ", files_sort[6])
            months.ford_2_completing()
        if choose == '8':
            print("You selected: ", files_sort[7])
            months.ford_leasing_completing()
        if choose == '9':
            print("You selected: ", files_sort[8])
            months.farmtrac_completing()
        if choose == '10':
            print("You selected: ", files_sort[9])
            months.ursus_completing()
        if choose == '11':
            print("You selected: ", files_sort[10])
            months.tym_completing()
        if choose == '12':
            print("You selected: ", files_sort[11])
            months.unimog_completing()
        if choose == '13':
            print("You selected: ", files_sort[12])
            months.unimog_2_completing()
        if choose == '14':
            print("You selected: ", files_sort[13])
            months.noremat_completing()


def repeat():
    print("ASDASDASDASD")
    quit()
    '''
    while True:
        message = input("Enter to continue, q to quit")
        if message == "q":
            break
        t = DisplayMenu()
        t.choose_year()
    '''


t = DisplayMenu()
t.choose_year()
