import cv2 
import pytesseract

img = cv2.imread('D:\web_scrap\captcha\thresh_img.jpg')

# Adding custom options
custom_config = r'--oem 3 --psm 6'
pytesseract.image_to_string(img, config=custom_config)