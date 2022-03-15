import os
import subprocess
import glob
import re
from natsort import natsorted
import start

class DisplayMenu:
    def years_sort(self):
        path = os.getcwd()
        '''
        for fileName_relative in glob.glob(path + "**/20**", recursive=True):
            fileName_absolute = os.path.basename(fileName_relative)
            print(fileName_absolute)
        '''

        folders = ([name for name in os.listdir(path) if os.path.isdir(os.path.join(path, name))])
        folders.reverse()
        p = re.compile(r".*(venv|.idea)")
        print([p.sub("", x) for x in folders])

    def choose_year(self):
        choose = input("Choose year 1.2022, 2.2023: ")
        if choose == "1":
            os.chdir("/home/btn/PycharmProjects/pojazdy/2022")
            t.cars_sort()
        if choose == "2":
            os.chdir("/home/btn/PycharmProjects/pojazdy/2023")
            t.cars_sort()


    def cars_sort(self):
        path = (os.getcwd())
        print(path)
        files = [f for f in glob.glob(path + "**/*.xlsx", recursive=True)]
        p = re.compile(r".*(/2022/|/2023/)")
        s = natsorted([p.sub('', x).strip() for x in files])
        print(s)
        #choose = input("Wybierz numer pojazdu: ")
        #if choose == '1':
        #    print(s[0])
        t.choose_car()


    def choose_car(self):
        os.chdir("/home/btn/PycharmProjects/pojazdy/")
        print(os.getcwd())
        #os.system('start.py')
        start.asd()


t = DisplayMenu()
#t.years_sort()                                                                      #OK
t.choose_year()
#t.cars_sort()                                                                      #OK
#t.choose_car()

##################################TO_DO
#



'''
    REGEX
        p = re.compile(r'.*(venv|.idea)')
        print([p.sub('', x).strip() for x in folders])
        print(folders)
    
      USTAWIANIE WYLISTOWANIA TYLKO CYFR/LAT
      char_seq = "1234567890"
      for char in char_seq:
          esc_set = "*" + glob.escape(char) + "*"
          for file in glob.glob(esc_set):
              print(file)



year = input("Wybierz rok: ")
print(int(year))
if year == '2022':
    path = os.getcwd()
    cars = [name for name in os.listdir(path) if not os.path.isdir(os.path.join(path, name))]
    print(cars)
    print(os.getcwd()) #sprawdzenie w jakim katalogu sie aktualnie znajdujemy
    os.chdir(path)     #wchodzimy do katalogu 'path'
    #f = open(cars[1])
    #print(f.read())
    #f.close()
    print(os.getcwd())
    #subprocess.call(['libreoffice', 'citroen.xlsx'])

else:
    path = "/home/btn/PycharmProjects/pojazdy/2023/"
    cars = [name for name in os.listdir(path) if not os.path.isdir(os.path.join(path, name))]
    print([name for name in os.listdir(path) if not os.path.isdir(os.path.join(path, name))])
    print(cars)
    print(os.getcwd()) #sprawdzenie w jakim katalogu sie aktualnie znajdujemy
    os.chdir(path)     #wchodzimy do katalogu 'path'
    f = open(cars[0])
    print(f.read())
    f.close()



LISTOWANIE PLIKOW Z ROZSZERZENIEM xlsx
# return all files as a list
for file in os.listdir(r"/home/btn/PycharmProjects/pojazdy/"):
    # check the files which are end with specific extension
    if file.endswith(".xlsx"):
        # print path name of selected files
        print([os.path.join(file)])


REGEX
lst =  ['Paris, 458 boulevard Saint-Germain', 'Marseille, 29 rue Camille Desmoulins', 'Marseille, 1 chemin des Aubagnens']
p = re.compile(r'.*(?:rue|boulevard|quai|chemin)')
print([p.sub('', x).strip() for x in lst])

'''
