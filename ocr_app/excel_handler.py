import pandas as pd
from tkinter import filedialog, messagebox

def extrac_column_row_excel(df: pd.DataFrame):
    
    try:
        excel_data = pd.read_excel(df)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file: {e}")
        return None, None
    
    # get columns and rows
    columns = excel_data.columns.tolist()
    rows = excel_data.values.tolist()
    
    # Tampilkan informasi kolom dan jumlah baris di messagebox
    messagebox.showinfo("Excel Data Loaded", f"Columns: {columns}\nRows loaded: {len(rows)}")
    
    return columns, rows


import pandas as pd
from tkinter import messagebox

def write_excel(sheet_name: str, columns, rows):
    if not sheet_name:
        messagebox.showinfo("No file selected", "Please select a file to save.")
        return
    
    try:
        # Cek apakah file sudah ada
        try:
            existing_data = pd.read_excel(sheet_name, sheet_name="Sheet1")
        except FileNotFoundError:
            # Jika file tidak ada, buat DataFrame kosong dengan kolom yang diinginkan
            existing_data = pd.DataFrame(columns=columns)
        
        # Buat DataFrame baru dari data yang akan ditambahkan
        new_data = pd.DataFrame(rows, columns=columns)
        
        # Gabungkan data lama dengan data baru
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        
        # Menulis data gabungan kembali ke file Excel dalam mode append
        with pd.ExcelWriter(sheet_name, mode="a", if_sheet_exists="overlay") as writer:
            updated_data.to_excel(writer, sheet_name="Sheet1", index=False)

        messagebox.showinfo("Excel Data Saved", f"File saved to: {sheet_name}")
    
    except Exception as e:
        messagebox.showerror("Error", f"Could not write file: {e}")
