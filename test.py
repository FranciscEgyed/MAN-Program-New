import os
import sqlite3
import pandas as pd
# Create your connection.

cnx = sqlite3.connect("F:\Python Projects\MAN 2022\MAN/Input/Others/database.db")

df = pd.read_sql_query("SELECT * FROM KSKDatabase", cnx)

print(df)