import csv
import os
from tkinter import filedialog, messagebox


def inlocuire_masterdata():
    indexro = ""
    array_ete_prelucrat = []
    path = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                   title="Selectati directorul cu fisiere:")
    pathsave = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                       title="Selectati directorul pentru salvare:")
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/ETE.txt", newline='') as csvfile:
        array_ete = list(csv.reader(csvfile, delimiter=';'))

    for i in range(len(array_ete[0])):
        array_ete_prelucrat.append([array_ete[0][i][0:10], array_ete[0][i][-1]])
    i = 0
    for (path, dirs, files) in os.walk(path):
        cale = path
        directoare = dirs
        i = + 1
        if i == 1:
            break

    for director in directoare:
        if not os.path.exists(pathsave + "/" + director):
            os.makedirs(pathsave + "/" + director)
        for file_all in os.listdir(path + "/" + director):
            partro = file_all[:-5].replace('23U', '23W')
            for i in range(len(array_ete_prelucrat)):
                if partro == array_ete_prelucrat[i][0]:
                    indexro = array_ete_prelucrat[i][1]
            partroindex = partro + indexro

            if file_all.endswith(".prg") and len(file_all) == 15:
                #path = os.path.join(path, file_all)
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read().replace(file_all[:-4], partroindex)
                with open(pathsave + "/" + director + "/" + partroindex + ".prg", "w") as fff:
                    fff.write(newText)

            elif file_all.endswith(".mdl"):
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read()
                with open(pathsave + "/" + director + "/" + partroindex + ".mdl", "w") as fff:
                    fff.write(newText)
            elif file_all.endswith(".csv"):
                with open(path + "/" + director + "/" + file_all, newline='') as csvfile:
                    csvarray = list(csv.reader(csvfile, delimiter=';'))
                for i in range(len(csvarray)):
                    for x in range(len(array_ete_prelucrat)):
                        for y, elem in enumerate(csvarray[i]):
                            if array_ete_prelucrat[x][0].replace("23W", "23U") in elem:
                                csvarray[i][y] = array_ete_prelucrat[x][0] + array_ete_prelucrat[x][1]
                with open(pathsave + "/" + director + "/" + file_all, 'w', newline='',
                          encoding='utf-8') as myfile:
                    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                    wr.writerows(csvarray)
            elif file_all.endswith(".CSV"):
                with open(path + "/" + director + "/" + file_all, newline='') as csvfile:
                    csvarray = list(csv.reader(csvfile, delimiter=';'))
                for i in range(len(csvarray)):
                    for x in range(len(array_ete_prelucrat)):
                        for y, elem in enumerate(csvarray[i]):
                            if array_ete_prelucrat[x][0].replace("23W", "23U") in elem:
                                csvarray[i][y] = array_ete_prelucrat[x][0] + array_ete_prelucrat[x][1]
                with open(pathsave + "/" + director + "/" + file_all, 'w', newline='',
                          encoding='utf-8') as myfile:
                    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                    wr.writerows(csvarray)
    messagebox.showinfo('Finalizat!')

