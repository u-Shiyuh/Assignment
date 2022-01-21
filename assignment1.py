from re import S
import numpy as np
import fitz
from pdf2image import convert_from_path
import glob, sys
import cv2
import pytesseract
import xlsxwriter
import os


path = os.chdir(os.path.dirname(__file__))
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# To get better resolution
zoom_x = 2.0  # horizontal zoom
zoom_y = 2.0  # vertical zoom
mat = fitz.Matrix(zoom_x, zoom_y)  # zoom factor 2 in each dimension

all_files_pdf = glob.glob("*.pdf")

#convert all files to png for opencv

for filename in all_files_pdf:
    doc = fitz.open(filename)  # open document
    for page in doc:  # iterate through the pages
        pix = page.get_pixmap(matrix=mat)  # render page to an image
        pix.save("page-%i.png" % page.number)  # store image as a PNG

#read all png 
all_files_png = glob.glob("*.png")

#initialize xlsx writer
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
row = 0

for img in all_files_png:
  
    img = cv2.imread(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    ret, thresh1 = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)
    rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 18))
    dilation = cv2.dilate(thresh1, rect_kernel, iterations = 1)
    contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL,
                                                    cv2.CHAIN_APPROX_NONE)
    im2 = img.copy()

    list_of_rect = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        list_of_rect.append(list([x, y, w, h]))

    #sort according to x coordinate
    list_of_rect = sorted(list_of_rect , key=lambda k: [k[0]])

    #to fix small difference in x coordinate
    sorted_list = []
    index = 0
    x_value = list_of_rect[0][0]

    for coor in list_of_rect:
        
        if coor[0] in range(x_value - 10, x_value + 10):
            coor[0] = x_value
        else:
            x_value = coor[0]
        sorted_list.append(coor)
    
    sorted_list = sorted(sorted_list , key=lambda k: [k[0], k[1]])


        
    for x,y,w,h in sorted_list:
        
        # Drawing a rectangle on copied image
        rect = cv2.rectangle(im2, (x, y), (x + w, y + h), (0, 255, 0), 2)
        
        # Cropping the text block for giving input to OCR
        cropped = im2[y:y + h, x:x + w]

        # Apply OCR on the cropped image
        text = pytesseract.image_to_string(cropped)
        if text != '':
            worksheet.write(row, 0 , text)
            row += 1



workbook.close()
print('end')
