import pandas as pd
from pymongo import MongoClient

# MongoDB connection setup
try:
    client = MongoClient('mongodb://localhost:27017/')
    db = client['ExcelData']  # The database 'ExcelData'
    print("MongoDB connection established successfully.")
except Exception as e:
    print(f"Failed to connect to MongoDB: {e}")

# Load Excel file dynamically with multiple sheets
try:
    excel_file_path = './ExcelFiles/file1.xlsx'
    excel_data = pd.read_excel(excel_file_path, sheet_name=None)  # Load all sheets
    print(f"Excel file loaded successfully. Sheets: {list(excel_data.keys())}")
except Exception as e:
    print(f"Failed to load Excel file: {e}")

# Iterate through each sheet and insert into separate MongoDB collections
for sheet_name, data in excel_data.items():
    # Convert all columns that are dates to handle NaT values
    for col in data.columns:
        if pd.api.types.is_datetime64_any_dtype(data[col]):
            # Convert datetime columns to string to avoid NaT issues
            data[col] = pd.to_datetime(data[col], errors='coerce').astype(str).replace('NaT', None)
        else:
            # For non-datetime columns, just replace NaN with None
            data[col] = data[col].where(pd.notnull(data[col]), None)

    # Convert data to dictionary format to insert into MongoDB
    data_dict = data.to_dict(orient='records')

    # Create collection for each sheet based on its name
    collection = db[sheet_name.replace(" ", "_")]  # MongoDB does not allow spaces in collection names

    # Insert data into MongoDB collection
    if data_dict:  # Only insert if the sheet contains data
        collection.insert_many(data_dict)
        print(f"Inserted data from sheet: {sheet_name} into collection: {sheet_name.replace(' ', '_')}")
    else:
        print(f"Sheet {sheet_name} is empty, skipping...")

print("Data import complete.")
