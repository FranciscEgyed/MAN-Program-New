import tkinter as tk
from PIL import ImageTk, Image
from functii_input import *
from diverse import structura_directoare, file_checker
from functii_prelucrare import *


def statusbusy():
    statuslabel["text"] = "Working on it . . . "


def statusidle():
    statuslabel["text"] = "Finished."


root = tk.Tk()
root.title("2022 MAN file processor")
root.geometry("450x450+50+50")
root.iconbitmap("Img/ICON.ico")
img = ImageTk.PhotoImage(Image.open("Img/MAN.jpg"))
container = tk.Frame(root, bg="black")
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
container = tk.Frame(root)
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
label = tk.Label(container, image=img)
menu_frame = tk.Frame(container)
statuslabel = tk.Label(root, text="Waiting . . .")

menu1 = tk.Menubutton(menu_frame, text="Fisiere Sursa", background="grey", font=("Arial", 10))
menu1.grid(row=0, column=0)
submenu1 = tk.Menu(menu1, tearoff=0, background="grey", font=("Arial", 10))
submenu1.add_command(label="Incarcare Input File", command=lambda: [statusbusy(), load_source(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix CSR", command=lambda: [statusbusy(), cmcsr(), statusidle()])
submenu1.add_command(label="Control Matrix CSL", command=lambda: [statusbusy(), cmcsl(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM L", command=lambda: [statusbusy(), cmtglml(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM R", command=lambda: [statusbusy(), cmtglmr(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL L(8025)", command=lambda: [statusbusy(), cm4axell(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL R(8026)", command=lambda: [statusbusy(), cm4axelr(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix Super Sleeve - in development",
                     command=lambda: [statusbusy(), cmss(), statusidle()])
menu1.configure(menu=submenu1)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu2 = tk.Menubutton(menu_frame, text="Fisiere Input")
menu2.grid(row=0, column=1)
submenu2 = tk.Menu(menu2, tearoff=0)
submenu2.add_command(label="Sortare JIT(SAP)", command=lambda: [statusbusy(), sortare_jit(), statusidle()])
submenu2.add_command(label="Sortare JIT(SAP) din director",
                     command=lambda: [statusbusy(), sortare_jit_dir(), statusidle()])
submenu2.add_command(label="Golire directoare LHD si RHD",
                     command=lambda: [statusbusy(), golire_directoare_comparati(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Prelucrare BOM-uri", command=lambda: [statusbusy(), boms(), statusidle()])
submenu2.add_command(label="Prelucrare WIRELIST-uri", command=lambda: [statusbusy(), wires(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Prelucrare BOM-uri cu PN Leoni",
                     command=lambda: [statusbusy(), boms_leoni(), statusidle()])
submenu2.add_command(label="Prelucrare WIRELIST-uri cu PN Leoni",
                     command=lambda: [statusbusy(), wires_leoni(), statusidle()])
menu2.configure(menu=submenu2)










menu_frame.grid(row=0, column=0)
container.grid(row=1, column=0)
label.grid(row=2, column=0)
statuslabel.grid(row=12, column=0)
structura_directoare()
file_checker()
root.mainloop()
