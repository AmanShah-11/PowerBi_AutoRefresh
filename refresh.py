import datetime
from datetime import datetime
import pbixrefresher
import time
import os
import sys
import argparse
import shutil
import psutil
# import pywintypes
# import win32api
from pywinauto.application import Application
from pywinauto import timings
import selenium
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import json

x = datetime.today()

# Dictionary to convert current day and month into the corresponding option on the drop-down menu
dict = {
    "1": '01', "2": '02', "3": '03', "4": '04', "5": '05', "6": '06', "7": '07', "8": '08', "9": '09', "10": '10',
    "11": '11', "12": '12', "13": '13', "14": '14', "15": '15', "16": '16', "17": '17', "18": '18', "19": '19',
    "20": '20', "21": '21', "22": '22', "23": '23', "24": '24', "25": '25', "26": '26', "27": '27', "28": '28',
    "29": '29', "30": '30', "31": '31'
}

dict_year = {
    "2011": "11", "2012": "12", "2013": "13", "2014": "14", "2015": "15", "2016": "16", "2017": "17", "2018": "18",
    "2019": "19", "2020": "20", "2021": "21", "2022": "22"
}

dict_month = {
    "1": "Jan",
    "2": "Feb",
    "3": "Mar",
    "4": "Apr",
    "5": "May",
    "6": "Jun",
    "7": "Jul",
    "8": "Aug",
    "9": "Sep",
    "10": "Oct",
    "11": "Nov",
    "12": "Dec",
}

# Allows typed words to be sent to powerbi window
def type_keys(string, element):
    """Type a string char by char to Element window"""
    for char in string:
        element.type_keys(char)


# Main function of program
def main():
    # Parse arguments from cmd
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", '--G:\\Total Rewards\\HR Metrics\\2018\\Power BI\\HR Metric Report - 2020.pbix', help="Path to .pbix file ")
    parser.add_argument("--workspace", help="name of online Power BI service work space to publish in",
                        default="My workspace")
    parser.add_argument("--refresh-timeout", help="refresh timeout", default=30000, type=int)
    parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default=True,
                        action='store_false')
    parser.add_argument("--init-wait", help="initial wait time on startup", default=90, type=int)

    args = parser.parse_args()
    # args.workbook = '--G:\\Total Rewards\\HR Metrics\\2018\\Power BI\\HR Metric Report - Simple Version.pbix'
    args.workbook = 'G:\\Total Rewards\\HR Metrics\\2018\\Power BI\\HR Metric Report - 2020.pbix'
    timings.after_clickinput_wait = 1

    WORKBOOK = args.workbook
    WORKSPACE = args.workspace
    INIT_WAIT = args.init_wait
    REFRESH_TIMEOUT = args.refresh_timeout

    # Kill running PBI
    PROCNAME = "PBIDesktop.exe"
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == PROCNAME:
            proc.kill()
    time.sleep(3)

    # Start PBI and open the workbook
    print("Starting Power BI")
    os.system('start "" "' + WORKBOOK + '"')
    print("Waiting ", INIT_WAIT, "sec")
    time.sleep(INIT_WAIT)
    print(WORKBOOK)
    filepath = WORKBOOK
    os.startfile(WORKBOOK)

    # Connect pywinauto
    print("Identifying Power BI window")
    app = Application(backend='uia').connect(path=PROCNAME)
    win = app.window(title_re='.*Power BI Desktop')
    time.sleep(60)
    win.wait("enabled", timeout=300)
    win.Save.wait("enabled", timeout=300)
    win.set_focus()
    win.Home.click_input()
    win.Save.wait("enabled", timeout=300)
    win.wait("enabled", timeout=300)
    # Refresh
    print("Refreshing")
    win.Refresh.click_input()
    # wait_win_ready(win)
    time.sleep(5)
    print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT, "sec)")
    win.wait("enabled", timeout=REFRESH_TIMEOUT)

    # Save
    # G:\Total Rewards\HR Metrics\2018\Material for Power BI
    # Use this path file to save the applause file.xlsx
    print("Saving")
    type_keys("%1", win)
    # wait_win_ready(win)
    time.sleep(5)
    win.wait("enabled", timeout=REFRESH_TIMEOUT)

    # Publish
    if args.publish:
        print("Publish")
        win.Publish.click_input()
        publish_dialog = win.child_window(auto_id="KoPublishToGroupDialog")
        publish_dialog.child_window(title=WORKSPACE).click_input()
        publish_dialog.Select.click()
        try:
            win.Replace.wait('visible', timeout=10)
        except Exception:
            pass
        if win.Replace.exists():
            win.Replace.click_input()
        win["Got it"].wait('visible', timeout=REFRESH_TIMEOUT)
        win["Got it"].click_input()

    # Close
    print("Exiting")
    win.close()

    # Force close
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()


# Loads browser and brings user to the website
def load_browser():
    url = "https://www.globoforce.net/microsites/t/home?client=caasco&setCAG=true"
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    driver = wd.Chrome(options=chrome_options)
    driver.get(url)
    return driver


# Signs into the applause website
def sign_in(driver):
    # Opens the secrets.txt file and obtains username and password
    try:
        with open('secrets.txt') as f:
            d = json.loads(f.read())
            usernameinput = d['username']
            passwordinput = d['password']
        # Finds the elements on the website for username, password, and signin button
        username = driver.find_element_by_name("username")
        password = driver.find_element_by_name("password")
        sign_in = driver.find_element_by_id("signIn-button")
        # Sends user info to username and password
        username.send_keys(usernameinput)
        password.send_keys(passwordinput)
        # Clicks sign in button
        sign_in.click()
    except Exception as e:
        print(e)


# Navigates to the tab containing reports on applause website
def navigate_to_report(driver):
    # Report tab used for testing purposes
    # # reports = driver.find_element_by_xpath('//a[@href="/microsites/t/awards/Redeem?client=caasco"]')
    # Report tab used for shane's applause (Actual)
    reports = driver.find_element_by_xpath('//a[@href="/microsites/t/reporting/ReportingHome"]')
    reports.click()


# Function to choose what day to start report from
def start_day(driver):
    press_button = driver.find_element_by_css_selector("div.selectize-control.a-select-control.searchParam.single")
    press_button.click()
    select_option = driver.find_element_by_xpath("//div[contains(text(),'01')]")
    # select_option.location_once_scrolled_into_view
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to choose what month to start report from
def start_month(driver):
    press_button = driver.find_elements_by_css_selector("div.selectize-control.a-select-control.searchParam.single")[1]
    press_button.click()
    select_option = driver.find_element_by_xpath("//div[contains(text(), 'Jan')]")
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to choose what year to start report from
def start_year(driver):

    press_button = driver.find_elements_by_css_selector("div.selectize-control.a-select-control.searchParam.single")[2]
    press_button.click()
    select_option = driver.find_element_by_xpath("//div[contains(text(), '2018')]")
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to choose what day to end report from
def end_day(driver):

    press_button = driver.find_elements_by_css_selector("div.selectize-control.a-select-control.searchParam.single")[3]
    press_button.click()
    select_option = press_button.find_element_by_xpath(".//div[contains(text(), '" + str(x.day) + "')]")
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to choose what month to end report from
def end_month(driver):

    press_button = driver.find_elements_by_css_selector("div.selectize-control.a-select-control.searchParam.single")[4]
    press_button.click()
    select_option = press_button.find_element_by_xpath(".//div[contains(text(),'" + dict_month[str(x.month)] + "')]")
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to choose what year to end report from
def end_year(driver):
    press_button = driver.find_elements_by_css_selector("div.selectize-control.a-select-control.searchParam.single")[5]
    press_button.click()
    select_option = press_button.find_element_by_xpath(".//div[contains(text(), '" + str(x.year) + "')]")
    driver.execute_script('arguments[0].scrollIntoView(true);', select_option)
    select_option.click()


# Function to download all the information from the report on applause website
def download_report(driver):
    select = driver.find_element_by_css_selector("button.a-btn.a-btn--secondary")
    select.click()


def move_old_file(path_file, file_name, path_file_move):
    dir_move = os.path.join(path_file_move)
    if not os.path.exists(dir_move):
        os.mkdir(path_file_move)
    shutil.move(os.path.join(path_file + file_name), path_file_move)


def move_new_file(download_path, path_file_move, file_name):
    try:
        old_path = os.path.join(download_path, file_name)
    except Exception as e:
        print(e)
        old_path = ""
    shutil.move(os.path.join(old_path, file_name), path_file_move)


# Calls all the functions to download report from functions
def applause():
    driver = load_browser()
    sign_in(driver)
    time.sleep(10)
    navigate_to_report(driver)
    time.sleep(10)
    start_day(driver)
    time.sleep(2)
    start_month(driver)
    time.sleep(2)
    start_year(driver)
    time.sleep(2)
    end_day(driver)
    time.sleep(2)
    end_month(driver)
    time.sleep(2)
    end_year(driver)
    time.sleep(2)
    download_report(driver)


# Starts the program(Runs program starting from here)
if __name__ == '__main__':
    try:
        applause()
        time.sleep(180)
        main()
        # time.sleep(20)
        # main()
    except Exception as e:
        print(e)




