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
# This code defines a function named `inlocuire_masterdata()`. The function uses the `filedialog` and `os` modules to
# prompt the user to select a directory containing files, and a directory to save the processed files.
# It then reads data from a file named `ETE.txt` and creates a processed version of this data in a list named
# `array_ete_prelucrat`.
# The function then loops through all directories in the selected directory, and for each directory, creates a new
# directory in the save directory with the same name. It then loops through all files in each directory and processes
# them based on their file extension. If a file has a `.prg` extension and a length of 15 characters, the function
# replaces the first 11 characters of the file name with a new string created by concatenating the first 10 characters
# of the original name with a single character read from `array_ete_prelucrat`. If a file has a `.mdl` extension,
# it copies the file to the new directory without making any changes. If a file has a `.csv` or `.CSV` extension,
# the function reads the file as a CSV, replaces any substrings matching certain patterns with new strings created
# by concatenating the matching substrings with a single character read from `array_ete_prelucrat`,
# and saves the updated CSV to the new directory.
# Finally, the function displays a message box indicating that the process is complete.
# Note: This code includes several commented-out lines that may have been used for testing or debugging, but are not
# currently being used.
