import os
import re
import csv
import pyautogui
import cv2 as cv
import pytesseract
import pandas as pd
from matplotlib import pyplot
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException

f = open("D:\web_scrap\\result\marks.csv", "w")
f.truncate()
f.close()

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation'])
option.add_experimental_option("detach",True)

browser = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=option)
browser.minimize_window()  
browser.set_window_size(600, 800)
browser.switch_to.window(browser.current_window_handle)

def check_success(err_code):
    print(err_code)

def fillLoginpage(usn, subject_codes, result_link):

    try:
        browser.get(result_link)
    except:
        return 404;
    
    textbox = browser.find_element(by=By.XPATH, value="/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
    captchabox = browser.find_element(by=By.XPATH, value="/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input")
    button = browser.find_element(by=By.XPATH, value="//*[@id='submit']")

    #time.sleep(1)
    myScreenshot = pyautogui.screenshot(region=(45, 415, 170, 80)) #region=(horizontal pos, vertical pos, vertical ratio, horizontal ratio)
    myScreenshot.save(r'D:\web_scrap\captcha\captcha_img.png') #change according to your dir.

    os.chdir('D:\web_scrap')
    img = cv.imread('captcha\captcha_img.png',0)
    ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
    os.chdir('D:\web_scrap\captcha')
    cv.imwrite("threshold_img.png", thresh)

    #time.sleep(1)
    img2 = cv.imread('threshold_img.png',0)
    #install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 
    custom_config = r'--oem 3 --psm 6'
    pre_captcha = pytesseract.image_to_string(img2, config=custom_config)
    pre_captcha.replace(" ", "").strip()
    captcha = re.sub('[^A-Za-z0-9]+', '', pre_captcha)
    #print("Printing solved Captcha " +captcha)

    if(len(captcha) != 6 ):
        return -1
    #time.sleep(1)

    try:
        textbox.send_keys(usn)
        captchabox.send_keys(captcha) 
        button.click()
    except:
        return -1

    try:
        obj = browser.switch_to.alert
        msg=obj.text
        obj.accept() #will click on ok
        #print(msg)
        if(msg == "Invalid captcha code !!!"):
            return -1
        if(msg == "University Seat Number is not available or Invalid..!"):
            return 1
    except NoAlertPresentException: 

        marks_list = []
        marks_list.append(usn)
        
        try:
            sub_code = 0
            while sub_code < len(subject_codes):
                internal_marks = browser.find_element(by=By.XPATH, value="//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[2]").text
                external_marks = browser.find_element(by=By.XPATH, value="//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[3]").text
                total_marks = browser.find_element(by=By.XPATH, value="//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[4]").text
                remarks = browser.find_element(by=By.XPATH, value="//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[5]").text

                marks_list.append(internal_marks)
                marks_list.append(external_marks)
                marks_list.append(total_marks)
                marks_list.append(remarks)
                sub_code +=1
        
        except NoSuchElementException:
            #print("Invalid (USN, Subject Code) combination")
            with open('D:\web_scrap\\result\marks.csv', 'a',) as f:
                write = csv.writer(f)
                marks_list.append("NA")
                write.writerow(marks_list)
            csv_read = pd.read_csv('D:\web_scrap\\result\marks.csv')
            csv2excel = pd.ExcelWriter('D:\web_scrap\\result\student_marks.xlsx')
            csv_read.to_excel(csv2excel)
            csv2excel.save()
            return 1 
        
        #print(marks_list)
        with open('D:\web_scrap\\result\marks.csv', 'a') as f:
            write = csv.writer(f)
            write.writerow(marks_list)
        csv_read = pd.read_csv('D:\web_scrap\\result\marks.csv')
        csv2excel = pd.ExcelWriter('D:\web_scrap\\result\student_marks.xlsx')
        csv_read.to_excel(csv2excel)
        csv2excel.save()

def main():
    ite=0
    resultLinkfile = open("D:\web_scrap\input\link.txt", "r")
    result_link = resultLinkfile.readline()
    #print(result_link)
    resultLinkfile.close()

    student_usn = []
    file = open('D:\web_scrap\\input\student_usn.csv')
    csvreader = csv.reader(file)
    for usns in csvreader:
        student_usn.append(usns[0])

    headerList = ['USN']
    subject_codes = []
    file = open('D:\web_scrap\\input\codes.csv')
    csvreader = csv.reader(file)
    for row in csvreader:
        subject_codes.append(row[0])
        headerList.append('Internals')
        headerList.append('Externals')
        headerList.append('Total')
        headerList.append('Remarks')
    #print(subject_codes)

    with open('D:\web_scrap\\result\marks.csv', 'a') as file:
        dw = csv.DictWriter(file, delimiter=',', fieldnames=headerList)
        dw.writeheader()

    while ite < 1:
        usn = student_usn[ite]
        #print(usn)
        x = fillLoginpage(usn, subject_codes, result_link)
        if(x == 1):
            ite = ite+1
            continue
        elif(x == -1):
            continue
        elif(x == 404):
            check_success(x)
        ite = ite+1
    check_success(1)

if __name__ == "__main__":
    main()