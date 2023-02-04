import gspread
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime


def login():
    driver.get("https://www.iotah-view.com/#/admin-panel/api-page")  # Go to IoTAh website
    driver.find_element(By.NAME, "email").send_keys("email")  # Find "email" and type in email
    driver.find_element(By.NAME, "password").send_keys("password")  # Find "password" and type in password
    driver.find_element(By.XPATH, "//button[@class='btn btn-primary btn-block rounded']").click()  # Click login button


def getAPI():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "url")))
    search = driver.find_element(By.NAME, "url")  # look for search bar
    search.click()  # click search bar
    search.send_keys("device/getAllSitesDevicesListing")  # type the api in search bar
    driver.find_element(By.XPATH, "//button[@class='btn btn-primary']").click()  # Find and click "send" button

    time.sleep(2)  # TODO: figure out way for program to work the instant API text appears
    file = json.loads(driver.find_element(By.XPATH, "//div[@class='mt-4 col-10 result']").text)  # Convert text to json
    return file


def filterList(devices):
    list = []

    for i in devices['devices']:  # filter out all devices from test sites and add the rest to a new list
        if i['customer_name'] != "SCT" and i['customer_name'] != "SCT Customer" \
                and i['customer_name'] != "Burris Refrigerated Logistics" \
                and i['customer_name'] != "SCT Demo Account" and i['customer_name'] != "EMS Returns - DONT TOUCH" \
                and i['customer_name'] != "C&S Grocers":
            list.append(i)

    return list


#  NOTE: Putting new devices in only works when there are no gaps, but there shouldn't be any in the first place
def updateSpreadsheet(new_list):
    currentTime = datetime.datetime.now()  # the time every device was updated

    account = gspread.service_account(filename='warnings.json')  # Accessing google account connected to spreadsheet
    sheet = account.open('IoTAhs with Warnings').sheet1  # Opening spreadsheet
    # NOTE: Sheets are organized by order not name (e.g. "Devices" = sheet1, "Total Warnings" = sheet2)

    for i in new_list:  # Going through every device in the list
        cell = sheet.find(i["mac_address"])  # Search for device in spreadsheet

        if cell is None:  # If device is not already in the spreadsheet, add the device at end of list
            row_counter = int(nextAvailableRow(sheet))
            updateDevice(sheet, row_counter, i, str(currentTime))
        else:  # If device exists in the spreadsheet, update its warning counters
            row_counter = cell.row
            getTime = sheet.cell(row_counter, 14).value  # get the last time device was updated
            lastTime = datetime.datetime.strptime(getTime, "%Y-%m-%d %H:%M:%S")  # convert getTime to correct data type

            if currentTime - lastTime >= datetime.timedelta(hours=24):  # only update device if it has been >= 24 hours
                updateDevice(sheet, row_counter, i, str(currentTime))
            time.sleep(2)


def nextAvailableRow(worksheet):  # Find the first empty row after the end of the list
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)


def updateDevice(sheet, row_counter, i, currentTime):  # Update info about the specific device (i)
    warnings = ["voltage_error_value", "rtc", "mis_voltage", "mis_capacity", "voltage_error_calibration",
                "current_sensor_open", "flash_size", "lost_rtc", "long_event"]

    #                      (row, column, value)
    sheet.update_cell(row_counter, 1, i["mac_address"])  # mac address
    sheet.update_cell(row_counter, 2, i["serial_number"])  # serial number
    sheet.update_cell(row_counter, 3, i["site_name"])  # site name
    sheet.update_cell(row_counter, 4, i["customer_name"])  # customer name
    sheet.update_cell(row_counter, 14, currentTime)  # time device was updated

    time.sleep(4)

    column_counter = 5  # Used for organizing which column the warning counter for each warning goes to
    for x in warnings:  # Check which warnings the device has
        if x in i["warnings"]:
            checkVal(sheet, row_counter, column_counter)
        else:
            clearIfEmpty(sheet, row_counter, column_counter)
        column_counter += 1

    row_counter += 1
    time.sleep(5)


def checkVal(sheet, row_counter, column):  # check if there is a counter already in the cell and if not, put 1 in cell
    val = sheet.cell(row_counter, column).value
    if val is None:
        sheet.update_cell(row_counter, column, 1)
    else:
        val = int(val)
        sheet.update_cell(row_counter, column, val + 1)


def clearIfEmpty(sheet, row_counter, column):  # if there is a counter in the cell but no warning, clear the cell
    val = sheet.cell(row_counter, column).value
    if val is not None:
        sheet.update_cell(row_counter, column, "")


driver = webdriver.Firefox(executable_path="C:/Users/npiaz/Downloads/geckodriver.exe")  # Accessing firefox driver
login()  # log into IoTAh website
devices = getAPI()  # get json file of API
driver.close()  # close firefox driver
new_list = filterList(devices)  # make list with test sites filtered out
updateSpreadsheet(new_list)  # post devices in spreadsheet
