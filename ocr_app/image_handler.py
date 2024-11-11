import cv2
from PIL import Image, ImageTk

def open_image(image_path):
    # Baca gambar dari path yang dipilih
    return cv2.imread(image_path)

def display_image(image, canvas):
    # Konversi gambar dari OpenCV (BGR) ke PIL (RGB) dan tampilkan di Tkinter canvas
    rgb_image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    pil_image = Image.fromarray(rgb_image)
    tk_image = ImageTk.PhotoImage(pil_image)
    canvas.config(width=tk_image.width(), height=tk_image.height())
    canvas.create_image(0, 0, anchor="nw", image=tk_image)
    return tk_image  # return reference agar gambar tidak dihapus
