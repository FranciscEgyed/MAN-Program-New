import tkinter as tk


def xxxx(data):
    def select_all():
        listbox.selection_set(0, tk.END)

    def clear_selection():
        listbox.selection_clear(0, tk.END)

    def print_selected():
        selected_indices = listbox.curselection()
        selected_items = [data[index][0] for index in selected_indices]
        print("Selected items:", selected_items)

    root = tk.Tk()
    root.title("Selectable List Example")
    listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)
    listbox.pack(fill=tk.BOTH, expand=True)
    for item, _ in data:
        listbox.insert(tk.END, item)

    select_all_button = tk.Button(root, text="Select All", command=select_all)
    select_all_button.pack(pady=5)

    clear_selection_button = tk.Button(root, text="Clear Selection", command=clear_selection)
    clear_selection_button.pack(pady=5)

    print_selected_button = tk.Button(root, text="Process Selected", command=print_selected)
    print_selected_button.pack(pady=5)

    description_label = tk.Label(root, text="", wraplength=300)
    description_label.pack(padx=10, pady=5)

    root.mainloop()


data = [('ID_CON355', 30), ('id_316_49', 105), ('Connector_occurrence_334', 8), ('id_316_2', 4),
        ('id_316_3', 4), ('id_316_4', 2)
        ]

sdsds = xxxx(data)
print()