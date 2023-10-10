for twist in lista_twisturi:
    index = ""
    for line in output:
        if twist[0] == line[34] and twist[1] not in lista_multicores:
            index = output.index(line) + 1
            indexfir1 = index - 4
            indexfir2 = index - 2
            twistwpa = ["", "", "", "", "", "", output[indexfir1][6] + "/" + output[indexfir2][6],
                        output[indexfir1][7] + "/" + output[indexfir2][7], "", "", "", "", "Twist WPA",
                        line[13], line[14], line[15], "", "", "", "", "", "", line[22], "", "", "", "",
                        "", "", line[29], "", "", "", "", line[34]]
    output.insert(index, twistwpa)
