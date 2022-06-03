from cgitb import text
from distutils.log import error
from lib2to3.pgen2 import driver
from tkinter.tix import Tree
from unicodedata import name
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import os
import pyautogui
import cv2 as cv
import pytesseract
from lxml import html
import requests

# install all these as pip install filename, and pip install opencv-python.

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_argument("start-maximized")
option.add_experimental_option("excludeSwitches", ['enable-automation'])
option.add_experimental_option("detach",True)

#add your chrome driver installation path
browser = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=option)

def fillLoginpage():

    browser.get("https://results.vtu.ac.in/FMEcbcs22/resultpage.php")

    #getting hold of usn and captcha input fields.

    testbox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
    captchabox = browser.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input")

    #start with the image capta recognition procedure
    time.sleep(2)
    myScreenshot = pyautogui.screenshot(region=(970, 405, 170, 80)) #region=(horizontal pos, vertical pos, vertical ratio, horizontal ratio)
    myScreenshot.save(r'D:\web_scrap\captcha\captcha_img.png') #change according to your dir.

    os.chdir('D:\web_scrap')
    img = cv.imread('captcha\captcha_img.png',0)
    ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
    #cv.imshow('Binary Threshold', thresh)
    # Using cv2.imwrite() method
    # Saving the image
    os.chdir('D:\web_scrap\captcha')
    cv.imwrite("thresh_img.png", thresh)

    time.sleep(1)
    #os.system('"wsl tesseract thresh_img.jpg result"') #tesseract is ocr function for image to text
    img2 = cv.imread('thresh_img.png',0)
    #install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 
    #depends on your tesseract installation path
    custom_config = r'--oem 3 --psm 6'
    captcha = pytesseract.image_to_string(img2, config=custom_config)
    
    captcha.replace(" ", "").strip()

    print("Captcha printing " +captcha)
    print(len(captcha)-1)
    if(len(captcha)-1 != 6 ):
        fillLoginpage()

    #finally input the result pages with required info.
    time.sleep(1)
    try:
        testbox.send_keys("1AM18CS077")
        captchabox.send_keys(captcha) 
    except:
        error
    try:
        print(browser.current_url)
    except:
        fillLoginpage()
    
    #copy the full XPATH for the required cell and add the below code to get the data
    subject_1 = browser.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[3]/div[2]").text
    total_1 = browser.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[3]/div[5]").text

    print(subject_1)
    print(total_1)
    
    time.sleep(100)

fillLoginpage()

#text = webdriver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[5]').text
#text = browser.find_element_by_xpath("/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]").text
#text = webdriver.find_element(By.XPATH,'//*[@id="raj"]/div[1]/div/label').text
#text = webdriver.find_element_by_xpath("//*[@id='dataPrint']/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[5]").text
#child_txt = driver.find_element_by_xpath("//div[@class='predictionsList']//div[@class='betWrapper ']//div[@class='betHeaderTitle']/span[@class='market']").text