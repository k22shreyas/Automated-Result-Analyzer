import time
import os
import pyautogui
import cv2 as cv
import pytesseract
#pip install all the above libraries

time.sleep(4)
myScreenshot = pyautogui.screenshot(region=(995,390, 170, 80)) #region=(vertical pos, horizontal pos, vertical ratio, horizontal ratio)
myScreenshot.save(r'D:\web_scrap\screenshot.png') #change according to your dir

os.chdir('D:\web_scrap')
img = cv.imread('screenshot.png',0)
ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
cv.imshow('Binary Threshold', thresh)
# Using cv2.imwrite() method
# Saving the image
os.chdir('D:\web_scrap\captcha')
cv.imwrite("thresh_img.jpg", thresh)

time.sleep(2)
#os.system('"wsl tesseract thresh_img.jpg result"') #tesseract is ocr function for image to text
img2 = cv.imread('thresh_img.jpg',0)
#install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 
#depends on your tesseract installation path
custom_config = r'--oem 3 --psm 6'
captcha = pytesseract.image_to_string(img2, config=custom_config)
print(captcha)
#captcha contains converted text