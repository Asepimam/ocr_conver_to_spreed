import tkinter as tk
from tkinter import filedialog
from image_handler import open_image, display_image
from ocr_handler import extract_text_from_boxes
from file_handler import save_to_new_excel,upload_excel
from excel_handler import write_excel

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Manual OCR with Red Boxes")
        
        # Frame untuk tombol
        button_frame = tk.Frame(root)
        button_frame.pack(side="top", fill="x")

        # Tombol untuk membuka gambar
        open_button = tk.Button(button_frame, text="Open Image", command=self.open_image)
        open_button.pack(side="left")

        # Tombol untuk melakukan OCR
        ocr_button = tk.Button(button_frame, text="Extract Text", command=self.extract_text)
        ocr_button.pack(side="left")

        # Tombol untuk menyimpan ke Excel
        save_button = tk.Button(button_frame, text="Save to Excel", command=self.save_to_excel)
        save_button.pack(side="left")

        # Tombol untuk mengunggah dan membaca file Excel
        upload_button = tk.Button(button_frame, text="Upload Excel", command=self.upload_excel_handler)
        upload_button.pack(side="left")
        
        # tombol hapus gambar
        clear_button = tk.Button(button_frame, text="Clear Image", command=self.clear_image)
        clear_button.pack(side="left")

        # Canvas untuk menampilkan gambar
        self.canvas = tk.Canvas(root, cursor="cross")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Frame untuk menampilkan kolom Excel dan kotak input
        self.data_frame = tk.Frame(root, padx=10, pady=10)
        self.data_frame.pack(side="right", fill="y")

        
        # Inisialisasi variabel
        self.image_path = None
        self.image = None
        self.red_boxes = []
        self.start_x = self.start_y = self.end_x = self.end_y = None
        self.text_data = []
        self.columns = []
        self.entry_fields = {}
        self.box_counter = 1  # Counter untuk memberi nomor pada kotak
        self.name_excel = None
        self.sheet_name = None

        # Event listener untuk menggambar kotak merah
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

    def open_image(self):
        # Buka dialog file untuk memilih gambar
        self.image_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if not self.image_path:
            return
        
        # Baca gambar dan tampilkan di canvas
        self.image = open_image(self.image_path)
        self.tk_image = display_image(self.image, self.canvas)

    def on_mouse_down(self, event):
        # Catat posisi awal dari klik mouse
        self.start_x, self.start_y = event.x, event.y

    def on_mouse_drag(self, event):
        # Saat mouse digerakkan, gambar kotak merah
        self.end_x, self.end_y = event.x, event.y
        self.canvas.delete("rect")
        self.canvas.create_rectangle(self.start_x, self.start_y, self.end_x, self.end_y, outline="red", width=2, tag="rect")

    def on_mouse_up(self, event):
        # Simpan koordinat kotak merah ke dalam list
        self.end_x, self.end_y = event.x, event.y
        self.red_boxes.append((self.start_x, self.start_y, self.end_x, self.end_y))
        
        # Tambahkan kotak merah dengan nomor urut di pojok kanan atas
        box_id = self.box_counter
        self.canvas.create_rectangle(self.start_x, self.start_y, self.end_x, self.end_y, outline="red", width=2)
        self.canvas.create_text(self.end_x - 10, self.start_y + 10, text=str(box_id), fill="red", font=("Arial", 10, "bold"))
        
        # Ekstrak teks OCR untuk kotak baru ini
        text = extract_text_from_boxes(self.image, [self.red_boxes[-1]])[0]
        self.text_data.append(text)
        
        # Tambahkan hasil OCR ke kolom input di kanan
        self.add_to_input_fields(text)
        
        # Increment box counter
        self.box_counter += 1

    def add_to_input_fields(self, text):
        # Masukkan hasil OCR ke dalam kolom input yang sudah ada dari hasil load Excel
        for idx, column in enumerate(self.columns):
            if idx < len(self.text_data):  # Cek agar sesuai dengan jumlah hasil OCR
                self.entry_fields[column].delete(0, tk.END)  # Hapus teks lama jika ada
                self.entry_fields[column].insert(0, self.text_data[idx])

    def extract_text(self):
        # Lakukan OCR pada semua area yang ditandai kotak merah
        self.text_data = extract_text_from_boxes(self.image, self.red_boxes)
        result_text = "\n".join(self.text_data)
        print("Extracted Text:\n", result_text)

    def save_to_excel(self):
        # Validasi jika melakukan upload excel maka ambil nama file excel lalu save pada file tersebut
        if self.name_excel:
            # Tambahkan nilai kosong jika `text_data` kurang dari jumlah `columns`
            adjusted_data = self.text_data + [None] * (len(self.columns) - len(self.text_data))
            print({"sheet_name": self.sheet_name, "columns": self.columns, "text_data": adjusted_data})
            
            # Panggil write_excel dengan data yang sudah disesuaikan
            write_excel(self.name_excel, self.columns, [adjusted_data])
        else:
            # Jika tidak ada file Excel yang di-upload, simpan sebagai file baru
            adjusted_data = self.text_data + [None] * (len(self.columns) - len(self.text_data))
            save_to_new_excel([adjusted_data], self.columns)


    def upload_excel_handler(self):
        # Memanggil fungsi upload_excel dari excel_handler
        self.columns, _,self.name_excel,self.sheet_name = upload_excel()

       
        if self.columns:
            self.display_excel_columns()

    def display_excel_columns(self):
        # Kosongkan frame data jika ada isinya sebelumnya
        for widget in self.data_frame.winfo_children():
            widget.destroy()

        # Tampilkan setiap kolom sebagai label dan buat kolom input otomatis untuk setiap kolom
        for idx, column in enumerate(self.columns):
            label = tk.Label(self.data_frame, text=column, anchor="w")
            label.grid(row=idx, column=0, sticky="w", padx=5, pady=5)

            entry = tk.Entry(self.data_frame)
            entry.grid(row=idx, column=1, padx=5, pady=5)

            # Simpan referensi input di dictionary untuk akses mudah
            self.entry_fields[column] = entry
            
    def clear_image(self):
        # Hapus gambar dan kotak merah dari canvas
        self.canvas.delete("all")
        self.image = None
        self.image_path = None
        self.red_boxes.clear()
        self.text_data.clear()
        self.box_counter = 1
        print("Image and all boxes have been cleared.")