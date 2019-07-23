from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import xlwt 
from xlwt import Workbook 
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QMainWindow, QPushButton, QLabel
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
import sys
import datetime
import pandas as pd
import numpy as np
import math
import os
import time

now = datetime.datetime.now()

username = ""
password = ""
#Need to move out

class App2(QMainWindow):
    def __init__(self, startDate, endDate, parent=None):
        super().__init__()
        self.title = 'Care Logs'
        self.left = 200
        self.top = 200
        self.width = 400
        self.height = 250
        self.startDate = startDate
        self.endDate = endDate
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.label = QLabel('Complete', self)
        self.label.resize(280,40)
        self.label.move(150,90)
        self.button = QPushButton('Begin Log Extraction', self)
        self.button.move(100,150)
        self.button.resize(200,40)

        self.label.hide()
        self.button.clicked.connect(lambda: self.collection(self.startDate, self.endDate))

#Completion Box

    @pyqtSlot()
    def collection(self, startDate, endDate):
        self.label.show()
        browser = webdriver.Chrome()
        browser.set_window_position(-10000,0)

        #browser.set_window_position(-10000,0)
        url = "https://goldmedal.clearcareonline.com/reports/custom/extreme/#/saved/carelogs-by-client/90547"
        browser.get(url)

        userTextbox = browser.find_element_by_id('id_username')
        passwordTextbox = browser.find_element_by_id('id_password')
        submitButton = browser.find_element_by_xpath("//input[@value='Login']")

        userTextbox.send_keys(username)
        passwordTextbox.send_keys(password)
        submitButton.click() 

        browser.get(url)

        browser.implicitly_wait(20)
        dateStart = browser.find_element_by_xpath("//div[@class='col-5 span5']/datepicker/input")
        dateStart.clear()
        dateStart.send_keys(self.startDate)

        dateEnd = browser.find_element_by_xpath("(//div[@class='col-5 span5']/datepicker/input)[2]")
        dateEnd.clear()
        dateEnd.send_keys(self.endDate)

        runReport = browser.find_element_by_xpath("//button[@class='btn btn-success']")
        runReport.click()

        try:
            saveReport = WebDriverWait(browser, 300).until(
            EC.element_to_be_clickable((
            By.XPATH, "//button[@title='Best option for preserving formatting of data and summary lines.']")))
        except TimeoutException:
            print("Loading took too much time!")

        saveReport.click()

#Selenium work

        time.sleep(10)
        browser.close()
        adder = 0.0
        iterator = 1
        authorizedHours = 0
        weeksLeft = 0
        hoursLeftPerWeek = 0
        hoursPerWeekInAuthorization = 0
        filename = " "

        directory = os.listdir(os.getcwd())

        for file in directory:
            if file.find("care_logs") > -1:
                filename = file
                break

        wb = Workbook() 
        sheet1 = wb.add_sheet('Sheet 1') 

        sheet1.write(0, 0, "Client (Last, First)")
        sheet1.write(0, 1, "Total Hours Used")
        sheet1.write(0, 2, "Authorized Hours Total")
        sheet1.write(0, 3, "Weeks Left in Authorization")
        sheet1.write(0, 4, "Authorized Hours Per Week")
        sheet1.write(0, 5, "Hours Left Per Week")

        data2018 = pd.read_csv("ADL data 2018.csv")
        data2019 = pd.read_csv("ADL data 2019.csv")
        careLogs = pd.read_excel(filename)

        print(data2018.columns.tolist())
        print(data2019.columns.tolist())

        names2018 = sorted(set(data2018['ClientNames']))
        names2019 = sorted(set(data2019['ClientNames']))
        namesCare = sorted(set(careLogs['Client Name']), reverse=True)

        df2018 = pd.DataFrame(data2018)
        df2019 = pd.DataFrame(data2019)
        dfCare = pd.DataFrame(careLogs)

#Read excel sheets
        for name in namesCare:
            frontName = name.split(" ")
            backName = frontName[1] + ", " + frontName[0]

            startDateExists = False
            startDateOff = None
            endDateOff = None
            authorizedAmountOff = None

            sheet1.write(iterator, 0, backName)
            print(name, " / ", backName)

            for row in dfCare.itertuples():
                if name == row[1]:
                    if not(pd.isnull(row[5])) and startDateExists == True:
                        newDate = datetime.datetime.strptime(str(row[5]), "%Y-%m-%d %H:%M:%S")
                        if startDateOff < newDate:
                            startDateOff = newDate
                            endDateOff = datetime.datetime.strptime(str(row[6]), "%Y-%m-%d %H:%M:%S")
                            authorizedAmountOff = row[3]

                    if not(pd.isnull(row[5])) and startDateExists == False:
                        startDateOff = datetime.datetime.strptime(str(row[5]), "%Y-%m-%d %H:%M:%S")
                        endDateOff = datetime.datetime.strptime(str(row[6]), "%Y-%m-%d %H:%M:%S")
                        authorizedAmountOff = row[3]
                        startDateExists = True


            if startDateExists == False:        
                startDateOff = datetime.datetime(2000, 6, 7)
                endDateOff = datetime.datetime.now()
                authorizedAmountOff = 0

            for row in dfCare.itertuples():
                if name == row[1]:
                    if not(pd.isnull(row[5])) and startDateOff == datetime.datetime.strptime(str(row[5]), "%Y-%m-%d %H:%M:%S"):
                        if not(pd.isnull(row[2])):
                            adder = adder + float(row[2])

            for row in df2018.itertuples():
                if backName == row[1]:
                    if startDateOff < datetime.datetime.strptime(row[3], "%m/%d/%Y %H:%M"):
                        print(startDateOff, " ", row[3])
                        adder = adder + float(row[24])

            for row in df2019.itertuples():
                if backName == row[1]:
                    if startDateOff < datetime.datetime.strptime(row[3], "%m/%d/%Y %H:%M"):
                        print(startDateOff, " ", row[3])
                        adder = adder + float(row[22])

            currentDate = datetime.datetime.now()
            weeksLeft = ((endDateOff - currentDate).days) / 7
            weeksTotal = ((endDateOff - startDateOff).days) / 7

            authorizedHours = weeksTotal * authorizedAmountOff
            hoursPerWeekInAuthorization = authorizedHours / weeksTotal

            if weeksLeft == 0:
                hoursLeftPerWeek = 0
            else:
                hoursLeftPerWeek = (authorizedHours - adder) / weeksLeft
            if authorizedHours == 0:
                authorizedHours = 0
                hoursLeftPerWeek = 0
                weeksLeft = 0
                hoursPerWeekInAuthorization = 0
#Calculation
            sheet1.write(iterator, 1, adder)
            sheet1.write(iterator, 2, authorizedHours)
            sheet1.write(iterator, 3, weeksLeft)
            sheet1.write(iterator, 4, hoursPerWeekInAuthorization)
            sheet1.write(iterator, 5, hoursLeftPerWeek)

            adder = 0.0
            iterator += 1

        wb.save('CareReport.xls')

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Care Logs'
        self.left = 200
        self.top = 200
        self.width = 400
        self.height = 250
        self.startDate = ""
        self.endDate = ""
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.label = QLabel('Enter Start Date', self)
        self.label.resize(280,40)
        self.label.move(20,0)

        # Start Date
        self.textbox = QLineEdit(self)
        self.textbox.move(20, 30)
        self.textbox.resize(280,40)

        self.label2 = QLabel('Enter End Date', self)
        self.label2.resize(280,40)
        self.label2.move(20,60)

        # End Date
        self.textbox2 = QLineEdit(self)
        self.textbox2.move(20, 90)
        self.textbox2.resize(280,40)

        self.button = QPushButton('Submit Dates', self)
        self.button.move(20,150)
        self.button.resize(150,40)

        self.button2 = QPushButton('Next', self)
        self.button2.move(200,150)
        self.button2.resize(150,40)

        self.label2 = QLabel('Saved', self)
        self.label2.resize(150 ,40)
        self.label2.move(20,180)

        self.label2.hide()

        self.show()
        # connect button to function on_click
        self.button.clicked.connect(self.on_click)
        self.button2.clicked.connect(self.next)


    @pyqtSlot()
    def on_click(self):
        self.startDate = self.textbox.text()
        self.endDate = self.textbox2.text()
        self.label2.show()

    @pyqtSlot()
    def next(self):
        self.w = App2(self.startDate, self.endDate)
        self.w.show()
        self.hide()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    app.exec_()
