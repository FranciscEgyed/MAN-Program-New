import os
import sqlite3
import pandas as pd
# Create your connection.

cnx = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")

df = pd.read_sql_query("SELECT * FROM KSKDatabase", cnx)

print(df)
pivot = df.pivot_table(index="Segment", columns="Part No", values="Cantitate", fill_value=0, aggfunc='count')
indexes = pivot.index.values.tolist()
valori = pivot.values.tolist()
coloane = pivot.columns.tolist()