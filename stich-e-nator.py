#!/usr/bin/env python
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait

from ctypes import windll
from PIL import Image

from selenium import webdriver
import urllib.request
import threading
import traceback
import requests
import keyboard
import win32gui
import zipfile
import os.path
import win32ui
import time
import os
import io
import re

class screenshot(object):
    def __init__(self,driver):
        self.driver = driver
        self.frames = 0 
        self.state = False
        self.previous_window_title = ""
    def capture(self,file_name,window_name):
        hwnd = win32gui.FindWindow(None, window_name)
        left, top, right, bot = win32gui.GetClientRect(hwnd)
        w = right - left      
        h = bot - top
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        im = Image.frombuffer('RGB',(bmpinfo['bmWidth'], bmpinfo['bmHeight']),bmpstr, 'raw', 'BGRX', 0, 1)
        if result == 1:
            im.save(file_name)
            win32gui.DeleteObject(saveBitMap.GetHandle()) 
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            return True
        return False    
    def capture_increment(self,window_name):
        if(os.path.isdir("frames") == False):
          os.mkdir("frames")
        if(self.capture("frames/%08d.png" % self.frames, window_name) == True):
            self.frames +=1
            return True
        return False
    def reset_increment(self):
        self.frames = 0
    def thread(self):
     while True:
        if self.state:
            self.driver.switch_to.window(driver.window_handles[0])
            if (self.driver.title == self.previous_window_title):
              if self.driver.title == "Discord":
                    formatted_title = driver.title
              else:
                    formatted_title = "%s - Discord" % self.driver.title
              if self.driver.execute_script("return document.querySelectorAll(\'[class*=\"scrollerBase\"]\')[2].scrollTop") == 0:
                  self.capture_increment(formatted_title)
                  print("Done! Stopping screenshots")
                  self.reset_increment()
                  self.state = False
                  continue
              self.capture_increment(formatted_title)
              self.driver.execute_script("document.querySelectorAll(\'[class*=\"scrollerBase\"]\')[2].scrollTop -= 350")
            else:
                print("Stopping screenshot, window changed")
                self.reset_increment()
                self.state = False
            self.previous_window_title = self.driver.title
        time.sleep(0.5)


if __name__ == "__main__":
    #Get current discord version
    discord_path = os.getenv('LOCALAPPDATA') + "\\Discord\\"
    for contents in os.listdir(discord_path):
     if "app-" in contents:
         print("Version: %s" % ( contents[4:-1]))
         app = contents
    discord_path = discord_path + app + "\\Discord.exe"
    #Setup selenium
    options = webdriver.ChromeOptions()
    options.binary_location = discord_path
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        if os.path.exists("chromedriver.exe"):
          driver = webdriver.Chrome(executable_path=".\\chromedriver.exe", options=options)
        else:  
          driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
    except Exception as e: 
        process_version = re.search("(((((\d+\.)(\d+\.))(\d+\.))(\*|\d+)))",str(e))
        if(process_version):
            process_version = process_version.group()
            print("Process version: %s" % (process_version))
            process_version_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" + process_version.rsplit(".",1)[0]
            print("Downloading version from: %s" % (process_version_url))
            contents = urllib.request.urlopen(process_version_url).read().decode("utf-8")
            process_url = "https://chromedriver.storage.googleapis.com/" + contents + '/chromedriver_win32.zip'
            print("Downloading chromedriver.exe %s from %s" % (contents,process_url))
            r = requests.get(process_url, allow_redirects=True)
            open('chromedriver_win32.zip', 'wb').write(r.content)
            with zipfile.ZipFile(io.BytesIO(r.content)) as zip_ref:
              zip_ref.extractall(".\\")
            print("Done, re-run this script")
            exit()
        else:
            traceback.print_exc()
            exit()
    WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(1))
    driver.switch_to.window(driver.window_handles[0])
    scr = screenshot(driver)
    t1 = threading.Thread(target=scr.thread, args=[])
    t1.daemon = True
    t1.start()
    try:
     while True:
        if keyboard.is_pressed('F7'):
            if scr.state == False:
                print("starting screenshot")
                driver.switch_to.window(driver.window_handles[0])
                scr.previous_window_title = driver.title
                scr.state = True
            else:
                print("screenshot already in progress")
        elif keyboard.is_pressed('F8'):
            if scr.state:
               print("stopping screenshot")
               scr.reset_increment()
               scr.state = False
            else:
                print("its already stopped")
        time.sleep(0.1)
    except KeyboardInterrupt:
        driver.quit()
        exit()