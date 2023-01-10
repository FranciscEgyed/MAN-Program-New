import csv
import os
from tkinter import filedialog, messagebox


def inlocuire():
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
            if file_all.endswith(".prg"):
                #path = os.path.join(path, file_all)
                partro = file_all[:-4].replace('23U', '23W')
                for i in range(len(array_ete_prelucrat)):
                    if partro == array_ete_prelucrat[i][0]:
                        indexro = array_ete_prelucrat[i][1]
                partroindex = partro + indexro
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read().replace(file_all[:-4], partroindex)
                with open(pathsave + "/" + director + "/" + partroindex + ".prg", "w") as fff:
                    fff.write(newText)
            elif file_all.endswith(".mdl"):
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read()
                with open(pathsave + "/" + director + "/" + file_all.replace("23U", "23W"), "w") as fff:
                    fff.write(newText)
            elif file_all.endswith(".csv"):
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read().replace("23U", "23W")
                with open(pathsave + "/" + director + "/" + file_all, "w") as fff:
                    fff.write(newText)
            elif file_all.endswith(".CSV"):
                with open(path + "/" + director + "/" + file_all) as f:
                    newText = f.read().replace("23U", "23W")
                with open(pathsave + "/" + director + "/" + file_all, "w") as fff:
                    fff.write(newText)
    messagebox.showinfo('Finalizat!')

