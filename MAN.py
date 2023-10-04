import tkinter as tk
from PIL import ImageTk, Image
from diagrame import *
from functii_crearediagrame import prelucrare_json, xmltojson, selectie_conectori
from functii_clustering import clustering
from functii_database import databasecontent, exportdatabase, database_delete_record, database_delete_all_records
from functii_diverse import *
from functii_eng import extragere_welding
from functii_etichete import *
from functii_input import *
from diverse import *
from functii_ksklight import *
from functii_pentruvalidare import wirelist_validare, bom_validare
from functii_prelucrare import *
from functii_prelucrare_ksk import *
from functii_prod import *
from functii_rapoarte import *
from masterdata import *


def statusbusy():
    statuslabel["text"] = "Working on it . . . "


def statusidle():
    statuslabel["text"] = "Finished."


root = tk.Tk()
root.title("2022 MAN file processor")
root.geometry("670x450+50+50")
root.iconbitmap("Img/ICON.ico")
img = ImageTk.PhotoImage(Image.open("Img/MAN.jpg"))
container = tk.Frame(root, bg="gray")
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
container = tk.Frame(root)
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
label = tk.Label(container, image=img)
menu_frame = tk.Frame(container)
statuslabel = tk.Label(root, text="Waiting . . .")

menu1 = tk.Menubutton(menu_frame, text="Fisiere Sursa", background="DarkSeaGreen1", font="Arial 10 bold")
menu1.grid(row=0, column=0)
submenu1 = tk.Menu(menu1, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu1.add_command(label="Incarcare Fisier Sursa", command=lambda: [statusbusy(), load_source(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix CSR(8013-8014)", command=lambda: [statusbusy(), cmcsrnew(), statusidle()])
submenu1.add_command(label="Control Matrix CSL(8011-8012)", command=lambda: [statusbusy(), cmcslnew(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM L(8023)", command=lambda: [statusbusy(), cmtglml(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM R(8024)", command=lambda: [statusbusy(), cmtglmr(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL L(8025)", command=lambda: [statusbusy(), cm4axell(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL R(8026)", command=lambda: [statusbusy(), cm4axelr(), statusidle()])
submenu1.add_command(label="Control Matrix ALL",
                     command=lambda: [statusbusy(), cmall(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix to EXCEL",
                     command=lambda: [statusbusy(), cmtoexcel(), statusidle()])

menu1.configure(menu=submenu1)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu2 = tk.Menubutton(menu_frame, text="Fisiere Input", background="DarkSeaGreen1", font="Arial 10 bold")
menu2.grid(row=0, column=1)
submenu2 = tk.Menu(menu2, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu2.add_command(label="Sortare JIT(SAP)", command=lambda: [statusbusy(), sortare_jit(), statusidle()])
submenu2.add_command(label="Sortare JIT(SAP) din director",
                     command=lambda: [statusbusy(), sortare_jit_dir(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Incarcare BOM-uri", command=lambda: [statusbusy(), boms(), statusidle()])
submenu2.add_command(label="Incarcare WIRELIST-uri", command=lambda: [statusbusy(), wires(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Creare Wirelist complet", command=lambda: [statusbusy(), wires_complet(), statusidle()])
submenu2.add_command(label="Creare Wirelist complet cu PN Leoni",
                     command=lambda: [statusbusy(), wires_pnleoni(), statusidle()])
submenu2.add_command(label="Creare BOM cu PN Leoni",
                     command=lambda: [statusbusy(), boms_pnleoni(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Creare BOM toate platformele",
                     command=lambda: [statusbusy(), boms_cumulat(), statusidle()])
submenu2.add_command(label="Creare Wirelist toate platformele",
                     command=lambda: [statusbusy(), wires_cumulat(), statusidle()])
menu2.configure(menu=submenu2)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu3 = tk.Menubutton(menu_frame, text="Prelucrare KSK", background="DarkSeaGreen1", font="Arial 10 bold")
menu3.grid(row=0, column=2)
submenu3 = tk.Menu(menu3, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu3.add_command(label="Wirelist individual",
                     command=lambda: [statusbusy(), wirelist_individual(), statusidle()])
submenu3.add_command(label="Wirelist toate", command=lambda: [statusbusy(), wirelist_director(), statusidle()])
submenu3.add_command(label="Wirelist all validare", command=lambda: [statusbusy(), wirelist_validare(), statusidle()])
submenu3.add_separator()
submenu3.add_command(label="BOM individual", command=lambda: [statusbusy(), prelucrare_individuala_bom(), statusidle()])
submenu3.add_command(label="BOM toate", command=lambda: [statusbusy(), bom_director(), statusidle()])
submenu3.add_command(label="BOM all validare", command=lambda: [statusbusy(), bom_validare(), statusidle()])
menu3.configure(menu=submenu3)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu4 = tk.Menubutton(menu_frame, text="Rapoarte", background="DarkSeaGreen1", font="Arial 10 bold")
menu4.grid(row=0, column=3)
submenu4 = tk.Menu(menu4, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu4.add_command(label="Raport", command=lambda: [statusbusy(), creare_raport(), statusidle()])
submenu4.add_command(label="Rapoarte din director",
                     command=lambda: [statusbusy(), creare_raport_director(), statusidle()])
submenu4.add_command(label="Rapoarte toate", command=lambda: [statusbusy(), creare_raport_all(), statusidle()])
menu4.configure(menu=submenu4)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu5 = tk.Menubutton(menu_frame, text="KSK Light", background="DarkSeaGreen1", font="Arial 10 bold")
menu5.grid(row=0, column=4)
submenu5 = tk.Menu(menu5, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu5.add_command(label="Raport KSK Light", command=lambda: [statusbusy(), raport_light(), statusidle()])
submenu5.add_separator()
submenu5.add_command(label="Lista taiere KSK Light", command=lambda: [statusbusy(), cutting_ksklight(), statusidle()])
submenu5.add_command(label="Lista SuperSleeve KSK Light", command=lambda: [statusbusy(), ss_ksklight(), statusidle()])
submenu5.add_separator()
menu5.configure(menu=submenu5)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++a
menu7 = tk.Menubutton(menu_frame, text="Diverse", background="DarkSeaGreen1", font="Arial 10 bold")
menu7.grid(row=0, column=5)
submenu7 = tk.Menu(menu7, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu7.add_command(label="Extragere lungimi KSK",
                     command=lambda: [statusbusy(), extragere_lungimi_ksk(), statusidle()])
submenu7.add_command(label="Extragere BOM KSK", command=lambda: [statusbusy(), extragere_bom_ksk(), statusidle()])
submenu7.add_command(label="Extragere Variatii de lungimi",
                     command=lambda: [statusbusy(), extragere_variatii(), statusidle()])
submenu7.add_separator()
submenu7.add_command(label="Stergere fisiere", command=golire_directoare, background="red")
submenu7.add_separator()
menu7.configure(menu=submenu7)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu8 = tk.Menubutton(menu_frame, text="Database", background="DarkSeaGreen1", font="Arial 10 bold")
menu8.grid(row=0, column=6)
submenu8 = tk.Menu(menu8, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu8.add_command(label="Extragere KSK din database",
                     command=lambda: [statusbusy(), databasecontent(), statusidle()])
submenu8.add_separator()
submenu8.add_command(label="Incarcare stock pentru comparatie", command=stockcompa)
submenu8.add_separator()
submenu8.add_command(label="Extragere database completa",
                     command=lambda: [statusbusy(), exportdatabase(), statusidle()])
submenu8.add_separator()
submenu8.add_command(label="Stergere inregistrari din database",
                     command=lambda: [statusbusy(), database_delete_record(), statusidle()], background="red")
submenu8.add_command(label="!!! Stergere database !!!",
                     command=lambda: [statusbusy(), database_delete_all_records(), statusidle()], background="red")
submenu8.add_separator()
menu8.configure(menu=submenu8)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu9 = tk.Menubutton(menu_frame, text="ENG / PROD", background="DarkSeaGreen3", font="Arial 10 bold")
menu9.grid(row=0, column=7)
submenu9 = tk.Menu(menu9, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenucascade5 = tk.Menu(submenu9, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenucascade6 = tk.Menu(submenu9, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")

submenu9.add_cascade(label="Functii Productie", menu=submenucascade5)
submenucascade5.add_command(label="APFR", command=lambda: [statusbusy(), apfr(), statusidle()])
submenucascade5.add_command(label="Big files breakup", command=lambda: [statusbusy(), breaklargefiles(), statusidle()])
submenu9.add_cascade(label="Functii ENG", menu=submenucascade6)
submenucascade6.add_command(label="Generare cod QR",
                     command=lambda: [statusbusy(), eticheteqr(), statusidle()])
submenucascade6.add_command(label="Prelucrare masterdata",
                     command=lambda: [statusbusy(), inlocuire_masterdata(), statusidle()])
submenucascade6.add_command(label="Comparatie fisiere",
                     command=lambda: [statusbusy(), comparatie_fisiere(), statusidle()])
submenucascade6.add_command(label="Clustering", command=lambda: [statusbusy(), clustering(), statusidle()])
submenucascade6.add_command(label="extragere Welding", command=lambda: [statusbusy(), extragere_welding(), statusidle()])

submenucascade6.add_separator()
submenucascade6.add_command(label="Comparatie diagrame", command=lambda: [statusbusy(),
                                                                         comparatiediagrame(), statusidle()])
submenucascade6.add_command(label="Extragere informatii din diagrame",
                           command=lambda: [statusbusy(), extragere_informatii_diagrame(), statusidle()])
submenucascade6.add_separator()
submenucascade6.add_command(label="Prelucrare fisiere Matrix Module",
                           command=lambda: [statusbusy(), crearematrixmodule(), statusidle()])
submenucascade6.add_command(label="++++Basic Module",
                           command=lambda: [statusbusy(), crearebasicmodule(), statusidle()])
submenucascade6.add_command(label="++++Lista diagrame in KSK",
                           command=lambda: [statusbusy(), diagrame_ksk(), statusidle()])
submenucascade6.add_separator()
submenucascade6.add_command(label="++++Faza 1 JSON din XML", command=lambda: [statusbusy(), xmltojson(), statusidle()])
submenucascade6.add_command(label="++++Faza 2 EXCEL din JSON", command=lambda: [statusbusy(), prelucrare_json(),
                                                                            statusidle()])
submenucascade6.add_command(label="++++Faza 3 ", command=lambda: [statusbusy(), selectie_conectori(),statusidle()])
menu9.configure(menu=submenu9)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

menu_frame.grid(row=0, column=0)
container.grid(row=1, column=0)
label.grid(row=2, column=0)
statuslabel.grid(row=12, column=0)
databasebackup()
structura_directoare()
file_checker()
databesemerge()
root.mainloop()

# This is a Python script that imports several modules and creates a GUI using the tkinter library.
# The GUI consists of a window with a title, an icon, and a size of 600x450 pixels. It contains a frame with an image,
# a menu bar with three drop-down menus (Fisiere Sursa, Fisiere Input, and Prelucrare KSK), and a status label.
# Each drop-down menu contains several options that can be clicked to perform different functions related to the
# corresponding category. The script also defines two functions, "statusbusy()" and "statusidle()", which update the
# text of the status label to indicate whether the program is working on a task or has finished.
