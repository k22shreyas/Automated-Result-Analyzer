import pyautogui
import os

myScreenshot = pyautogui.screenshot(region=(0,0, 300, 400))
myScreenshot.save(r'D:\screenshot.png')
os.system('wsl ~ -e sh -c " tesseract \\wsl.localhost\Ubuntu\home\thor\sum\captcha\\image1.png stdout"')