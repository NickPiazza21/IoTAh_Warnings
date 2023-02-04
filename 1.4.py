import xlwings
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

    ws = xlwings.Book('IoTAhs with Warnings.xlsx').sheets['Sheet1']

    for i in new_list:  # Going through every device in the list
        cell = find(i["mac_address"], new_list, ws)

        # If device is not already in the spreadsheet, add the device at end of list. Else, update its warning counters.
        if cell is None:
            row_counter = nextAvailableRow(ws)
            updateDevice(ws, row_counter, i, currentTime)
        else:
            row_counter = cell
            getTime = ws[row_counter, 13].value  # get the last time device was updated

            if currentTime - getTime >= datetime.timedelta(hours=24):  # only update device if it has been >= 24 hours
                updateDevice(ws, row_counter, i, currentTime)


def find(query, new_list, ws):  # Search for device in spreadsheet
    for i in range(1, len(new_list) + 1):
        if query == ws[i, 0].value:
            return i

    return None


def nextAvailableRow(ws):  # Find the first empty row at the end of the list
    cel = ws.range("A1:A2")
    rng = cel.current_region
    last_cel = rng.end("down")
    empty_cell = last_cel.offset(1, 0)
    empty_cell = str(empty_cell)
    empty_cell = empty_cell.replace("<Range [IoTAhs with Warnings.xlsx]Sheet1!$A$", "")
    empty_cell = empty_cell.replace(">", "")
    empty_cell = int(empty_cell) - 1

    return empty_cell


def updateDevice(ws, row_counter, i, currentTime):  # Update info about the specific device (i)
    warnings = ["voltage_error_value", "rtc", "mis_voltage", "mis_capacity", "voltage_error_calibration",
                "current_sensor_open", "flash_size", "lost_rtc", "long_event"]

    # (row, column)
    ws[row_counter, 0].value = i["mac_address"]  # mac address
    ws[row_counter, 1].value = i["serial_number"]  # serial number
    ws[row_counter, 2].value = i["site_name"]  # site name
    ws[row_counter, 3].value = i["customer_name"]  # customer name
    ws[row_counter, 13].value = currentTime  # time device was updated

    column_counter = 4  # Used for organizing which column the warning counter for each warning goes to
    for x in warnings:  # Check which warnings the device has
        if x in i["warnings"]:
            checkVal(ws, row_counter, column_counter)
        else:
            clearIfEmpty(ws, row_counter, column_counter)
        column_counter += 1

    row_counter += 1


def checkVal(ws, row_counter, column):  # check if there is a counter already in the cell and if not, put 1 in cell
    val = ws[row_counter, column].value
    if val is None:
        ws[row_counter, column].value = 1
    else:
        val = int(val)
        ws[row_counter, column].value = val + 1


def clearIfEmpty(ws, row_counter, column):  # if there is a counter in the cell but no warning, clear the cell
    val = ws[row_counter, column].value
    if val is not None:
        ws[row_counter, column].value = ''


driver = webdriver.Firefox(executable_path="C:/Users/npiaz/Downloads/geckodriver.exe")  # Accessing firefox driver
login()  # log into IoTAh website
devices = getAPI()  # get json file of API
driver.close()  # close firefox driver
new_list = filterList(devices)  # make list with test sites filtered out
updateSpreadsheet(new_list)  # post devices in spreadsheet
