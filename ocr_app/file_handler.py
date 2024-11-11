import pandas as pd
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
def save_to_new_excel(data, columns=None):
    """
    Simpan data hasil OCR ke file Excel baru. 
    - data: List dari baris yang akan disimpan.
    - columns: Daftar nama kolom yang akan digunakan dalam file Excel.
    """
    if not data:
        print("No text to save. Please extract text first.")
        return

    # Gunakan kolom default jika tidak diberikan
    if columns is None:
        columns = ["Extracted Text"]

    # Buat DataFrame dari data yang ada
    df = pd.DataFrame(data, columns=columns)
    
    # Dialog simpan sebagai file Excel baru
    output_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_excel:
        try:
            df.to_excel(output_excel, index=False)
            messagebox.showinfo("Success", f"Text saved to {output_excel}")
            print(f"Text saved to {output_excel}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")

    

import pandas as pd
from tkinter import filedialog, messagebox

def upload_excel():
    # Buka dialog untuk memilih file Excel
    input_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not input_excel:
        messagebox.showinfo("No file selected", "Please select an Excel file.")
        return None, None, None, None

    try:
        # Baca file Excel untuk mendapatkan daftar sheet
        excel_file = pd.ExcelFile(input_excel)
        sheet_names = excel_file.sheet_names  # Dapatkan semua nama sheet

        # Jika ada lebih dari satu sheet, tampilkan pilihan untuk memilih sheet
        if len(sheet_names) > 1:
            sheet_name = messagebox.askstring("Select Sheet", f"Available sheets: {sheet_names}\nPlease enter the sheet name:")
            if sheet_name not in sheet_names:
                messagebox.showerror("Invalid Sheet", "The sheet name you entered does not exist.")
                return None, None, None, None
        else:
            sheet_name = sheet_names[0]  # Jika hanya ada satu sheet, gunakan sheet tersebut

        # Baca data dari sheet yang dipilih
        excel_data = pd.read_excel(input_excel, sheet_name=sheet_name)

    except Exception as e:
        messagebox.showerror("Error", f"Could not read file: {e}")
        return None, None, None, None
    
    # Ambil kolom dan baris dari data Excel
    columns = excel_data.columns.tolist()
    rows = excel_data.values.tolist()
    
    # Tampilkan informasi kolom dan jumlah baris di messagebox
    messagebox.showinfo("Excel Data Loaded", f"Sheet: {sheet_name}\nColumns: {columns}\nRows loaded: {len(rows)}")
    
    return columns, rows, input_excel, sheet_name

