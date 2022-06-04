from cgitb import text
from distutils.log import error
from lib2to3.pgen2 import driver
from logging import exception
from textwrap import fill
from tkinter.tix import Tree
from unicodedata import name
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import openpyxl
import time
import os
import pyautogui
import pandas as pd
import cv2 as cv
import pytesseract


# install all these as pip install filename, and pip install opencv-python.
  

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_argument("start-maximized")
option.add_experimental_option("excludeSwitches", ['enable-automation'])
option.add_experimental_option("detach",True)

#add your chrome driver installation path
browser = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=option)


def fillLoginpage(usn, ite):

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
    #print(len(captcha)-1)
    if(len(captcha)-1 != 6 ):
        return -1

    #finally input the result pages with required info.
    time.sleep(1)

    try:
        testbox.send_keys(usn)
        captchabox.send_keys(captcha) 
    except:
        return -1
    try:
        print(browser.current_url)
    except:
        return -1
    
    time.sleep(2)
    sub_codes = ["18ME751", "18CS71", "18CS72","18CS744","18CS734","18CSL76","18CSP77"]
    rows = []

    for sub_code in sub_codes:
        subject = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[1]").text
        internal_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[2]").text
        external_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[3]").text
        total_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[4]").text
        remarks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+sub_code+"')]//following::div[5]").text

        present_row_data={'Subject Code': sub_code,
                   'Subject Name': subject,
                   'Internal Marks': internal_marks,
                   'External Marks': external_marks,
                   'Total': total_marks,
                   'Remarks': remarks }
        rows.append(present_row_data)
    
    final_result_data = pd.DataFrame(rows)                              #import pandas as pd
    final_result_data.to_excel(r'vtu_result.xlsx',index=False)
    wb.save("vtu_result.xlsx")

    time.sleep(2)
    return ite;


filepath=r"D:\web_scrap\score\student_marks_list.xlsx"    #excel path
wb=load_workbook(filepath)                                                         # load into wb
sheet=wb.active                                                                    # active workbook
#store and pass current usn to function
def main():
    ite=3
    print("START")
    while ite <= sheet.max_column:
        cell_obj = sheet.cell(row=ite, column=1)
        usn = cell_obj.value
        x = fillLoginpage(usn, ite)
        print("IN MAIN FUNC") #for testing
        print(ite) #for testing
        if(x == ite):
            print(x)
        elif(x == -1):
            print(x)
            continue
        ite = ite+1


if __name__ == "__main__":
    main()
