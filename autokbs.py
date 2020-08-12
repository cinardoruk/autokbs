'''
Copyright 2020 Cinar Doruk
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

'''

'''
TODO

*add a third button that does entries AND the removals afterwards.
    -select two csvs and load them into two separate dfs
    -file selection dialog boxes should have the appropriate titles
    -make the dentry, fentry, iterate, riterate functions accept dfs as arguments
    -
'''

UNAME=""
PASS=""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import UnexpectedAlertPresentException as alertception
from selenium.common.exceptions import NoSuchElementException        
from selenium.common.exceptions import TimeoutException

import tkinter as tk
from tkinter import filedialog
import pandas as pd
from time import sleep

global wait
wait = 5

def main():
    
    #bring up screen to ascertain whether it's guest entry or guest removal
    #the buttons are tied to the appropriate functions. see gui()
    #loadcsv()
    #siteAccess()
        
    gui()
    
#draw a little window with buttons for data entry and removal
def gui():
    window = tk.Tk()
    window.title("KBS Giris-Cikis Otomasyon Sistemi")
    window.geometry('250x75')
    lbl = tk.Label(window, text="Yapacaginiz islemi secin.\n")
    lbl.grid(column=1, row=0)
    
    def removal():
        loadcsv()
        siteAccess()
        riterate()
        window.destroy()
        driver.close()
        
    def entry():
        loadcsv()
        siteAccess()
        iterate()
        driver.close()
        
    btn = tk.Button(window,                            \
                 text="Misafir Giris",                 \
                 command=entry)
    btn2 = tk.Button(window,                           \
                  text="Misafir Cikis",                \
                  command=removal)
    
    btn.grid(column=1, row=2)
    btn2.grid(column=2, row=2)
    window.mainloop()

#starting webdriver, getting website, filling in uname and pass(stored in UNAME and PASS, first two sloc in script.
def siteAccess():    
    global driver
    driver = webdriver.Chrome()
    driver.get("https://kbs.egm.gov.tr/")
    
    webdriver.DesiredCapabilities.CHROME["unexpectedAlertBehavious"] = "accept"
    
    uname = UNAME
    pswd = PASS
    
    u = driver.find_element_by_name("txtkullaniciadi")
    u.clear()
    u.send_keys(uname)
    
    p = driver.find_element_by_name("txtsifre")
    p.clear()
    p.send_keys(pswd)
    
    #wait for user to enter captcha and click enter
    WebDriverWait(driver, wait + 150).until(EC.title_is("Ana Sayfa"))
    print("page change detected")
    
    #switching to the "add guest" tab/page
    konak = driver.find_element_by_xpath("//span[@id='lblkonaklayan']")
    konak.click()
    
#fn for removing people from the registry
def remove(rowNum):

    WebDriverWait(driver, wait).until(EC.visibility_of_element_located((By.NAME, "txtAraGecer")))

    #differentiate between domestic and foreing rows
    if (df.isnull().iloc[rowNum,0]):
        driver.find_element_by_name("txtAraGecer").send_keys(df.iloc[rowNum][1])
    elif (df.isnull().iloc[rowNum,1]):
        driver.find_element_by_name("txtAraKimlik").send_keys(df.iloc[rowNum][0])
    else:
        print("invalid excel file")
        return
        
    driver.find_element_by_name("btnSorgula").send_keys(Keys.SPACE)

    rembut = '//*[@id="grdkonaklayan_ctl02_Sil"]'
    button = '/html/body/div[3]/div[2]/div/div/div/div/div[4]/button[1]'
    
    try:
        WebDriverWait(driver, wait).until(EC.element_to_be_clickable((By.XPATH, rembut)))
        driver.find_element_by_xpath(rembut).click()
        confirm(button, ' ')
    except (TimeoutException, NoSuchElementException):
        print("doesn't exist")
        pass

#selecting excel file and importing as pandas df
def loadcsv():
    global df
    
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename()
    
    df = pd.read_excel(file_path, dtype=str)
    
#domestic data entry
def dentry(rowNum):
    if not (driver.find_element_by_name("txttcno").is_displayed()):
        driver.find_element_by_id("home-tab").click()#send_keys(Keys.SPACE)

    WebDriverWait(driver, wait).until(EC.visibility_of_element_located((By.NAME, "txttcno")))

    driver.find_element_by_name("txttcno").send_keys(df.iloc[rowNum]['TC KimlikNo'])
    driver.find_element_by_id("imgmernis").click()
    driver.find_element_by_id("txtVerilenOda").send_keys(df.iloc[rowNum]['VerilenOda'])

    driver.find_element_by_id("btnTurk").send_keys(Keys.SPACE)
    
    button = '/html/body/div[1]/div[2]/div/div/div/div/div[4]/button[1]'

    confirm(button, ' ')
    
#foreign data entry
def fentry(rowNum):

    if not (driver.find_element_by_name("txtPasaport").is_displayed()):
        #WebDriverWait(driver, wait).until(EC.element_to_be_clickable((By.NAME, "profile-tab")))
        driver.find_element_by_id("profile-tab").click()#send_keys(Keys.SPACE)
    
    #wait until the tab switches to foreign entry, i.e the element "txtPasaport" is interactable    
    
    WebDriverWait(driver, wait).until(EC.visibility_of_element_located((By.NAME, "txtPasaport")))
        
    driver.find_element_by_name("txtPasaport").send_keys(df.iloc[rowNum][1])
    driver.find_element_by_name("txtYAdi").send_keys(df.iloc[rowNum][2])
    driver.find_element_by_name("txtYSoyadi").send_keys(df.iloc[rowNum][3])
    driver.find_element_by_name("txtYDogum").send_keys(df.iloc[rowNum][8])
    driver.find_element_by_name("txtYDogumYeri").send_keys(df.iloc[rowNum][7])
    driver.find_element_by_name("txtYVerilenOda").send_keys(str(df.iloc[rowNum][10]))
    if df.iloc[rowNum][4] == "Erkek":
        Select(driver.find_element_by_id('drp_listCinsiyet')).select_by_value('1')
    else:
        Select(driver.find_element_by_id('drp_listCinsiyet')).select_by_value('2')
    Select(driver.find_element_by_name('drpUyrugu')).select_by_value(str(df.iloc[rowNum][11]))
    
    driver.find_element_by_name("btnYabanci").send_keys(Keys.SPACE)
    #driver.find_element_by_name("btnYabanci").click()    
    
    #there'll be no conf pop up if the room is empty. gotta handle that
    
    button = '/html/body/div[1]/div[2]/div/div/div/div/div[4]/button[1]'
    button2 = '/html/body/div[2]/div[2]/div/div/div/div/div[4]/button[1]'
    #/html/body/div[2]/div[2]/div/div/div/div/div[4]/button[1]

    confirm(button, button2)

#takes care of the confirmation popup and success alert.
#is an abomination but gets the job done
def confirm(bttnpath, bttn2path):

    try:
        WebDriverWait(driver, wait).until(EC.visibility_of_element_located((By.XPATH, bttnpath)))
        driver.find_element_by_xpath(bttnpath).send_keys(Keys.SPACE)
    except:
        if bttn2path != ' ':
            try:
                WebDriverWait(driver, wait).until(EC.visibility_of_element_located((By.XPATH, bttn2path)))
                driver.find_element_by_xpath(bttn2path).send_keys(Keys.SPACE)
            except:
                try:
                    WebDriverWait(driver, wait).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                except:
                    pass
        else:
            try:
                WebDriverWait(driver, wait).until(EC.alert_is_present())
                driver.switch_to.alert.accept()
            except:
                pass
    try:
        WebDriverWait(driver, wait).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    except:
        pass

#iterate over rows to do data entry
def iterate():
    
    for i in range(df.shape[0]):
        if df.isnull().iloc[i,0]:
            #do foreign data entry
            fentry(i)
        else:
            #do domestic data entry
            dentry(i)
    print("done. whew.")
    
#iterate over rows to do data removal
def riterate():
    for i in range(df.shape[0]):
        remove(i)
    print("done. whew.")
    
#calling main
main()
