import os
from tkinter import filedialog
import pandas as pd

dir_BOM = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
for file_all in os.listdir(dir_BOM):
    if file_all.endswith(".csv"):
        df = pd.read_csv(dir_BOM + "/" + file_all)
        print(df.to_string())
