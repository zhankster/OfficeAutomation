try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import argparse
import sys
from pathlib import Path

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
def ocr_core(filename):
    """
    This function will handle the core OCR processing of images.
    """
    text = pytesseract.image_to_string(Image.open(filename))  # We'll use Pillow's Image class to open the image and pytesseract to detect the string in the image
    return text

print(ocr_core(sys.argv[1]))
#pth = r'%s' % sys.argv[1]
#print(ocr_core(pth))

#directory = sys.argv[1]
#new_file = Path(directory)
#print(new_file)

