import tkinter as tk
import sqlite3

# Create a connection to the SQLite database
conn = sqlite3.connect("sample.db")
cursor = conn.cursor()

# Create a table if it doesn't exist
cursor.execute("""
    CREATE TABLE IF NOT EXISTS contacts (
        id INTEGER PRIMARY KEY,
        name TEXT,
        email TEXT
    )
""")
conn.commit()

# Function to insert data into the database
def insert_data():
    name = name_entry.get()
    email = email_entry.get()
    cursor.execute("INSERT INTO contacts (name, email) VALUES (?, ?)", (name, email))
    conn.commit()
    update_listbox()
    clear_entries()

# Function to update the listbox with data from the database
def update_listbox():
    listbox.delete(0, tk.END)
    cursor.execute("SELECT * FROM contacts")
    for row in cursor.fetchall():
        listbox.insert(tk.END, f"ID: {row[0]}, Name: {row[1]}, Email: {row[2]}")

# Function to clear the entry fields
def clear_entries():
    name_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)

# Create the main application window
root = tk.Tk()
root.title("Database Viewer and Editor")

# Create and configure widgets
name_label = tk.Label(root, text="Name:")
name_entry = tk.Entry(root)
email_label = tk.Label(root, text="Email:")
email_entry = tk.Entry(root)
insert_button = tk.Button(root, text="Insert", command=insert_data)
listbox = tk.Listbox(root)

# Grid layout
name_label.grid(row=0, column=0, padx=5, pady=5)
name_entry.grid(row=0, column=1, padx=5, pady=5)
email_label.grid(row=1, column=0, padx=5, pady=5)
email_entry.grid(row=1, column=1, padx=5, pady=5)
insert_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)
listbox.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

# Initial data population
update_listbox()

# Start the Tkinter main loop
root.mainloop()

# Close the database connection when the application is closed
conn.close()




