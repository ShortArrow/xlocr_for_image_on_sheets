from PIL import Image
import pytesseract

# If you don't have tesseract executable in your PATH, include the following:
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
# Example tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract'

image = Image.open('Untitled.jpg')
# Example of adding any additional options
custom_oem_psm_config = r'--psm 6 -l eng -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789"'
result = pytesseract.image_to_string(image, config=custom_oem_psm_config)

print(result)
