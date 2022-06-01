import time
import os
import pyautogui
import cv2 as cv

time.sleep(4)
myScreenshot = pyautogui.screenshot(region=(995,390, 170, 80)) #region=(vertical pos, horizontal pos, vertical ratio, horizontal ratio)
myScreenshot.save(r'D:\web_scrap\screenshot.png')

os.chdir('D:\web_scrap')
img = cv.imread('screenshot.png',0)
ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
cv.imshow('Binary Threshold', thresh)
# Using cv2.imwrite() method
# Saving the image
os.chdir('D:\web_scrap\captcha')
cv.imwrite("thresh_img.jpg", thresh)

time.sleep(2)
os.system('"wsl tesseract thresh_img.jpg result"') #tesseract is ocr function for image to text

