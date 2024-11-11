import cv2
import pytesseract



def extract_text_from_boxes(image, red_boxes):
    # Lakukan OCR pada area yang ditandai kotak merah
    text_data = []
    for (x1, y1, x2, y2) in red_boxes:
        # Crop area dari gambar berdasarkan kotak merah
        cropped_image = image[y1:y2, x1:x2]
        text = pytesseract.image_to_string(cropped_image, lang='eng')
        text_data.append(text.strip())
    return text_data


