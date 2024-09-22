import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style
from pymongo import MongoClient
import pandas as pd

# Function to handle the import process
def import_data(client_url, db_name, excel_file_path):
    try:
        # MongoDB connection setup
        client = MongoClient(client_url)

        existing_dbs = client.list_database_names()
        if db_name not in existing_dbs:
            db_name = db_name.replace(" ", "_")
            update_status(f"Database '{db_name}' not found. Creating a new database called '{db_name}'...")
            update_message(f"A database '{db_name}' has been created successfully.")

        db = client[db_name]

        # Load Excel file dynamically with multiple sheets
        update_status("Loading Excel file...")
        excel_data = pd.read_excel(excel_file_path, sheet_name=None)

        total_sheets = len(excel_data)
        sheet_count = 0

        for sheet_name, data in excel_data.items():
            sheet_count += 1
            update_status(f"Processing sheet {sheet_name} ({sheet_count}/{total_sheets})...")

            # Convert all columns that are dates to handle NaT values
            for col in data.columns:
                if pd.api.types.is_datetime64_any_dtype(data[col]):
                    data[col] = pd.to_datetime(data[col], errors='coerce').astype(str).replace('NaT', None)
                else:
                    data[col] = data[col].where(pd.notnull(data[col]), None)

            # Convert data to dictionary format to insert into MongoDB
            data_dict = data.to_dict(orient='records')

            # Create collection for each sheet based on its name
            collection = db[sheet_name.replace(" ", "_")]

            # Insert data into MongoDB collection
            if data_dict:
                collection.insert_many(data_dict)
                update_status(f"{sheet_name} imported successfully.")
            else:
                update_status(f"{sheet_name} is empty, skipping...")

            progress_var.set((sheet_count / total_sheets) * 100)
            root.update_idletasks()

        update_status(f"Data import complete! ({excel_file_path.split('/')[-1]} into {db_name})")
        update_message(f"{excel_file_path.split('/')[-1]} has been imported successfully into {db_name}.")
        messagebox.showinfo("Success", "Data import is complete.")

        # Reset progress bar and status after user clicks OK
        progress_var.set(0)
        update_status("Status: Waiting for a new import")

    except Exception as e:
        update_status(f"An error occurred: {e}")
        update_message(f"{excel_file_path.split('/')[-1]} import failed.")
        messagebox.showerror("Error", f"An error occurred: {e}")

        # Reset progress bar and status after failure
        progress_var.set(0)
        update_status("Status: Please try a new import")


# Function to update the status message
def update_status(message):
    status_text.insert(tk.END, message + "\n")
    status_text.yview(tk.END)  # Scroll to the end of the text area

def update_message(message):
    import_status_text.insert(tk.END, message + "\n")
    import_status_text.yview(tk.END)  # Scroll to the end of the text area

# Function to browse for Excel file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

# Function to start the import process
def start_import():
    client_url = client_entry.get()
    db_name = db_entry.get()
    excel_file_path = file_entry.get()

    if not client_url or not db_name or not excel_file_path:
        messagebox.showwarning("Missing information", "Please fill in all fields.")
        return

    import_data(client_url, db_name, excel_file_path)

# GUI setup
root = tk.Tk()
root.title("Excel to MongoDB Importer")
root.config(padx=20, pady=20)  # Add padding around the window

# Configure a modern style for widgets
style = Style()
style.configure("TButton", font=("Trebuchet MS", 12, "bold"), padding=10)
style.configure("TLabel", font=("Trebuchet MS", 12), padding=10)
style.configure("TEntry", font=("Trebuchet MS", 12), padding=10)
style.configure("TProgressbar", thickness=20)

# Labels and entries for MongoDB client URL, DB name, and Excel file path
tk.Label(root, text="MongoDB Client URL:", font=("Trebuchet MS", 12)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
client_entry = tk.Entry(root, width=40, font=12, relief="flat", borderwidth=2)
client_entry.grid(row=0, column=1, padx=10, pady=10, sticky="we")

tk.Label(root, text="Database Name:", font=("Trebuchet MS", 12)).grid(row=1, column=0, padx=10, pady=10, sticky="w")
db_entry = tk.Entry(root, width=40, font=12, relief="flat", borderwidth=2)
db_entry.grid(row=1, column=1, padx=10, pady=10, sticky="we")

tk.Label(root, text="Excel File:", font=("Trebuchet MS", 12)).grid(row=2, column=0, padx=10, pady=10, sticky="w")
file_entry = tk.Entry(root, width=40, font=12, relief="flat", borderwidth=2)
file_entry.grid(row=2, column=1, padx=10, pady=10, sticky="we")

tk.Button(root, text="Browse", command=browse_file, font=("Trebuchet MS", 12)).grid(row=2, column=2, padx=10, pady=10, sticky="we")

# Start import button
tk.Button(root, text="Start Import", command=start_import, font=("Trebuchet MS", 12)).grid(row=3, column=1, pady=10, sticky="we")

# Status text area
status_text = tk.Text(root, height=10, width=60, font=("Trebuchet MS", 12), relief="flat", borderwidth=2)
status_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Progress bar
progress_var = tk.DoubleVar()
progress_bar = Progressbar(root, variable=progress_var, maximum=100, style="TProgressbar")
progress_bar.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="we")

# Import status label to display success/failure
import_status_text = tk.Text(root, height=5, width=60, font=("Trebuchet MS", 12), relief="flat", borderwidth=2, fg="green")
import_status_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Configure the root grid rows and columns for expansion
root.grid_rowconfigure(4, weight=1)  # Make the status_text area expandable
root.grid_rowconfigure(6, weight=1)  # Make the import_status_text area expandable

root.grid_columnconfigure(1, weight=1)  # Make the center column expand with the window

# Run the GUI
root.mainloop()
