# Import every modul, what is necessary
import webbrowser
import os
import xlrd
import xlwt
from splinter import Browser
import pyautogui
from easygui import *
import sys
import datetime
import time
import pyprind

# Software START! I give an option to continue, or quit
image = "someImage.jpg"
msg = "Üdvözöllek a JunkJar 1.0 alkalmazásban!"
choices = ["Tovább","Kilépés"]
reply = buttonbox(msg, image=image, choices=choices)

if reply == "Kilépés":
        sys.exit()
else:
        pass

# Just a little reminder, that you'll have some tasks
msgbox('A következő lépésnél add meg a belépési adatokat, majd válassz ki egy excel fájlt, amit össze szeretnél vetni a honlap adatbázisával. Ha elkészültem, egy hibalistát fogsz találni ugyanott, ahonnan elindítottál.')

# Give me your username and password!
def dataCheck(data):
        if data == None:
                sys.exit()
        else:
                return data

# Check the correct datas. THIS IS NOT A CORRECT AUTHENTICATION FUNCTION, THE PAGE OWNS THE AUTHENTICATION!
loginFlag = False

while loginFlag == False:
        username = dataCheck(enterbox("Kérlek add meg a felhasználónevet"))
        password = dataCheck(passwordbox("Kérlek add meg a jelszót"))
        if username != "someUsername" or password != "somePassword":
                msgbox('Hibás bejelentkezési adatok!')
        else:
                loginFlag = True

# Choose the excel file!
filename = fileopenbox()

# Select datas from excel
w_row = 1
r_worksheet = xlrd.open_workbook(filename)
workbook = xlwt.Workbook(encoding = 'ascii')
w_worksheet = workbook.add_sheet('Hibák')
handle_sheet_0 = r_worksheet.sheet_by_index(0)

# Create date for excel
actualTime = datetime.datetime.now().strftime("%y-%m-%d-%H-%M")

# DON'T DO ANYTHING
image = "stop.jpg"
msg = "Kérlek ne csinálj semmit, ne nyomj meg semmit egészen addig, amíg el nem készül a folyamat, illetve magától be nem zárul minden ablak! Köszönöm!"
choices = ["Tovább","Kilépés"]
reply = buttonbox(msg, image=image, choices=choices)

if reply == "Kilépés":
        sys.exit()
else:
        pass

# Choose the browser (default is Firefox)
browser = Browser()

# Hide the browser a little
browser.driver.set_window_position(-10000,0)

# Fill in the url
browser.visit('somePageURL')

# Find the username cell
browser.find_by_name('user')[1].fill(username)

# Find the password cell
browser.find_by_name('password')[1].fill(password)

# Find the submit button and click
browser.find_by_css('.loginsub').first.click()

# If there is a "full sign", we kick somebody out :) Sorry!!
try:
        browser.find_by_tag('input')[3].click()

except:
        pass

# I setup a progress bar here
bar = pyprind.ProgBar(handle_sheet_0.nrows, stream=1)

# I create the excel header
style1 = xlwt.easyxf('pattern: pattern solid, fore_colour gray40; align: horiz center')
style2 = xlwt.easyxf('align: horiz center')
style3 = xlwt.easyxf('pattern: pattern solid, fore_colour red; align: horiz center')
style4 = xlwt.easyxf('pattern: pattern solid, fore_colour green; align: horiz center')

cellWidths = [75, #0 Rövid cégnév az excelben
              20, #1 Adószám az excelben
              20, #2 Irányítószám az excelben
              20, #3 Város az excelben
              75, #4 Cím az excelben
              75, #5 Rövid cégnév a honlapon
              75, #6 Teljes cégnév a honlapon
              25, #7 Adószám a honlapon
              25, #8 Teljes adószám a honlapon
              20, #9 Irányítószám a honlapon
              20, #10 Város a honlapon
              75, #11 Cím a honlapon
              25, #12 Cégjegyzék a honlapon
              25] #13 Státusz

for column in range(len(cellWidths)):
        w_worksheet.col(column).width = 256 * cellWidths[column]

headerLabels = ['Rövid cégnév az excelben',     #0, compare(0-5)
                'Adószám az excelben',          #1, compare(1-7)
                'Irányítószám az excelben',     #2, compare(2-9)
                'Város az excelben',            #3, compare(3-10)
                'Cím az excelben',              #4, compare(4-11)
                'Rövid cégnév a honlapon',      #5 
                'Teljes cégnév a honlapon',     #6
                'Adószám a honlapon',           #7
                'Teljes adószám a honlapon',    #8
                'Irányítószám a honlapon',      #9
                'Város a honlapon',             #10
                'Cím a honlapon',               #11
                'Cégjegyzék a honlapon',        #12
                'Státusz']                      #13

styleFlag = style1

for headerCell in range(len(headerLabels)):
        if headerCell > 4:
                styleFlag = style4
                
        w_worksheet.write(0, headerCell, headerLabels[headerCell], styleFlag)

# I setup the compare function
def compare(excelData, siteData, excelDataIndex, siteDataIndex):
        if excelData != siteData:
                w_worksheet.write(w_row, excelDataIndex, excelData, style3)
                w_worksheet.write(w_row, siteDataIndex, siteData, style2)
                global statusFlag
                statusFlag += 1 
        else:
                w_worksheet.write(w_row, excelDataIndex, excelData, style2)
                w_worksheet.write(w_row, siteDataIndex, siteData, style2)

# I setup the multiple line tester
def multipleLineTester(xpath, variableName):
        if '\n' in xpath:
                variableName = xpath[:xpath.index("\n")]
                return variableName
        else:
                variableName = xpath
                return variableName
        
# Checking the statuses
def statusChecker():
        if statusFlag == 5:
                w_worksheet.write(w_row, 13, 'Fatális hiba', style3)
        elif statusFlag >= 1:
                w_worksheet.write(w_row, 13, 'Hiba', style3)
        else:
                w_worksheet.write(w_row, 13, 'OK', style4)

def dataPull(valueName, siteDataPath, excelData, excelIndex):
        try:
                valueName = siteDataPath
                return 
        except:
                w_worksheet.write(w_row, excelIndex, excelData, style3)
                pass
                
   

# The program loops through the company names, and tax numbers
for row in range(handle_sheet_0.nrows):

        # Setting back the status flag
        statusFlag = 0

        # Update progress bar
        bar.update()

        # We give a loop flag, what is gonna throw an exception, if there is an absolutely bogus company name
        loop = 0

        # The examined company's name in the excel
        companyNameInExcel = handle_sheet_0.cell(row, 0).value

        # The examined company's tax number in the excel
        taxNumberInExcel = handle_sheet_0.cell(row, 1).value
        try:
                first8digitInExcel = int(str(taxNumberInExcel)[:8])
        except:
                pass

        # The examined company's zip code in the excel
        zipCodeInExcel = handle_sheet_0.cell(row, 2).value

        # The examined company's city in the excel
        cityInExcel = handle_sheet_0.cell(row, 3).value

        # The examined company's address in the excel
        addressInExcel = handle_sheet_0.cell(row, 4).value

        # To the search page!
        browser.visit('somePageURL')

        try:
        
                # Find the form and fill it with the data
                browser.find_by_name('adoszam').first.fill(first8digitInExcel)
                
                # Find the form and fill it with the data
                browser.click_link_by_id('send_button')

                # Press enter
                browser.find_by_id('cegnev0').click()
        except:
                pass

        # Wait until the element is loaded
        while browser.is_element_present_by_css('.cegnev') == False and loop < 5:

                loop += 1
                pass     
        try:
                # Storing the full name of the company
                companyFullNameValue = browser.find_by_xpath("//td[@class='contentsub'][starts-with(.,'2/')]/following-sibling::td[1]").last.value

                # Checking the values
                companyFullNameInSite = ''
                companyFullNameInSite = multipleLineTester(companyFullNameValue, companyFullNameInSite)

                # Writin the full name to the excel. We don't compare it to anything
                w_worksheet.write(w_row, 6, companyFullNameInSite, style2)

                # Storing the company's tax number
                taxNumberInSite = browser.find_by_id('ceg').value[-13:]
                w_worksheet.write(w_row, 8, taxNumberInSite, style2)
                first8digitInSite = int(str(taxNumberInSite)[:8])

                # Checking the values
                try:
                        compare(first8digitInExcel, first8digitInSite, 1, 7)
                except:
                        w_worksheet.write(w_row, 1, taxNumberInExcel, style3)
                        pass

                try:
                        # Storing the short name of the company
                        companyShortNameValue = browser.find_by_xpath("//td[@class='contentsub'][starts-with(.,'3/')]/following-sibling::td[1]").last.value
                        companyNameInSite = ''
                        companyNameInSite = multipleLineTester(companyShortNameValue, companyNameInSite)
                        compare(companyNameInExcel, companyNameInSite, 0, 5)
                except:
                        w_worksheet.write(w_row, 0, companyNameInExcel, style3)
                        statusFlag += 1
                        pass
                
        except:
                # Logs the error in the excel, if there was a problem with the company's name
                w_worksheet.write(w_row, 0, companyNameInExcel, style3)
                w_worksheet.write(w_row, 1, taxNumberInExcel, style3)
                w_worksheet.write(w_row, 2, zipCodeInExcel, style3)
                w_worksheet.write(w_row, 3, cityInExcel, style3)
                w_worksheet.write(w_row, 4, addressInExcel, style3)
                w_worksheet.write(w_row, 13, 'Fatális hiba', style3)
                w_row += 1
                continue

        # Storing the company's zip code, then check it
        zipCodeInSite = int(browser.find_by_css('.cim').first.value[0:4])
        compare(zipCodeInExcel, zipCodeInSite, 2, 9) 

        # Storing the company's city, then check it
        fullAddress = browser.find_by_css('.cim').first.value
        cityNamesEndingIndex = fullAddress.index(",")
        cityInSite = browser.find_by_css('.cim').first.value[5:cityNamesEndingIndex]
        compare(cityInExcel, cityInSite, 3, 10) 
                
        # Storing the company's address, then check it
        addressInSite = browser.find_by_css('.cim').first.value[(cityNamesEndingIndex+2):]
        compare(addressInExcel, addressInSite, 4, 11) 

        # Print the company's register number
        registerNumberInSite = browser.find_by_id('ceg').value[-37:-25]
        w_worksheet.write(w_row, 12, registerNumberInSite, style2)

        statusChecker()

        w_row += 1

# Close the excel, and every process
workbook.save('../Hibák/Hibalista-' + actualTime + '.xls')
os.system("taskkill /im firefox.exe /f")
os.system("taskkill /im geckodriver.exe /f")
