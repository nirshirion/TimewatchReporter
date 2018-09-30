from datetime import datetime
import argparse
from argparse import RawTextHelpFormatter
import os
import re
try:
    from selenium import webdriver
except Exception as e:
    print("Selenium not found - install it with the following command:\npip install selenium")
    os._exit(1)
try:
    from selenium.webdriver.chrome.options import Options
except Exception as e:
    print("Chromedriver not found - install it with the following command:\nbrew install chromedriver")
    os._exit(1)
from selenium.webdriver.support.wait import WebDriverWait

current_time = datetime.now()
current_year = current_time.year
current_month = current_time.month


def hour(s):
    if s not in ["{:02d}".format(x) for x in range(25)]:
        raise argparse.ArgumentTypeError("value has to be between 00 and 24")
    return s


def minutes(s):    
    if s not in ["{:02d}".format(x) for x in range(59)]:
        raise argparse.ArgumentTypeError("value has to be between 00 and 59")
    return s


def month(s):
    if not s.isdigit():
        raise argparse.ArgumentTypeError("value has to be a number between 1 and 12")
    m = int(s)
    if 1 <= m <= 12:
        return m
    raise argparse.ArgumentTypeError("value has to be a number between 1 and 12")


def year(s):
    min_year = current_year - 9
    max_year = current_year + 1
    if not s.isdigit():
        raise argparse.ArgumentTypeError("value has to be a number between {} and {}".format(min_year, max_year))
    y = int(s)
    if min_year <= y <= max_year:
        return y
    raise argparse.ArgumentTypeError("value has to be a number between {} and {}".format(min_year, max_year))

parser = argparse.ArgumentParser(description="Required dependencies:\npip install selenium\nbrew install chromedriver", formatter_class=RawTextHelpFormatter)
parser.add_argument("company_number")
parser.add_argument("employee_number")
parser.add_argument("password")
parser.add_argument("start_hour", nargs="?", default="09", type=hour, help="Format is hh, default is 09")
parser.add_argument("start_minutes", nargs="?", default="00", type=minutes, help="Format is mm, default is 00")
parser.add_argument("end_hour", nargs="?", default="18", type=hour, help="Format is hh, default is 18")
parser.add_argument("end_minutes", nargs="?", default="00", type=minutes, help="Format is mm, default is 00")
parser.add_argument("-y", "--year", nargs="?", default=str(current_year), type=year, help="Format is yyyy, default is current year")
parser.add_argument("-m", "--month", nargs="?", default=str(current_month), type=month, help="Format is m, default is current month")
parser.add_argument("-s", "--silent", action="store_true", help="Run the program in silent mode")
args = parser.parse_args()

company_number = args.company_number
employee_number = args.employee_number
password = args.password
start_hour = args.start_hour
start_minutes = args.start_minutes
end_hour = args.end_hour
end_minutes = args.end_minutes
weekend_days = {4, 5}
holiday_txts = ["חג", "ערב חג"]
silent = args.silent
custom_month = args.month
custom_year = args.year

options = Options()
if silent:
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
try:
    driver = webdriver.Chrome(options=options)
except Exception as e:
    print("Chromedriver not found - install it with the following command:\nbrew install chromedriver")
    os._exit(1)

driver.maximize_window()
driver.get("http://www.timewatch.co.il")
driver.find_element_by_xpath("//a[@href='http://checkin.timewatch.co.il/punch/punch.php']").click()
driver.find_element_by_id("compKeyboard").send_keys(company_number)
driver.find_element_by_id("nameKeyboard").send_keys(employee_number)
driver.find_element_by_id("pwKeyboard").send_keys(password)
driver.find_element_by_name("B1").submit()
if custom_month != current_month or custom_year != current_year:
    custom_url = driver.find_element_by_xpath("//a[starts-with(@href, '/punch/editwh.php')]").get_attribute("href")
    if custom_month != current_month:
        custom_url = re.sub(r"m=..", "m={:02d}".format(custom_month), custom_url)
    if custom_year != current_year:
        custom_url = re.sub(r"y=..", "y={}".format(custom_year), custom_url)
    driver.get(custom_url)
else:
    driver.find_element_by_xpath("//a[starts-with(@href, '/punch/editwh.php')]").click()
driver.find_element_by_xpath("//tr[@class='tr']/td[last()]").click()
WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) == 2)
driver.switch_to.window(driver.window_handles[1])
page_date = datetime.strptime(driver.find_element_by_name("d").get_attribute("value"), "%Y-%m-%d")
small_day_txt = driver.find_element_by_xpath("//form/table/tbody/tr[7]/td/table/tbody/tr/td[2]/font[2]/b").text
absence_reason = driver.find_element_by_xpath("//form/table/tbody/tr[8]//table/tbody/tr[8]//table//tr[1]//select/option[@selected]").get_attribute("value")
current_month = page_date.month
page_month = page_date.month
while current_month == page_month:
    if ((page_date.weekday() not in weekend_days) and
        (small_day_txt not in holiday_txts) and
        (absence_reason == "0")):
        print("Trying to update {} with {}:{}-{}:{}".format(page_date.strftime("%d-%m-%Y"), start_hour, start_minutes, end_hour, end_minutes))
        start_hour_txt = driver.find_element_by_id("ehh0")
        if start_hour_txt.get_attribute("value") == "":
            start_hour_txt.clear()
            start_hour_txt.send_keys(start_hour)
        start_minutes_txt = driver.find_element_by_id("emm0")
        if start_minutes_txt.get_attribute("value") == "":
            start_minutes_txt.clear()
            start_minutes_txt.send_keys(start_minutes)
        end_hour_txt = driver.find_element_by_id("xhh0")
        if end_hour_txt.get_attribute("value") == "":
            end_hour_txt.clear()
            end_hour_txt.send_keys(end_hour)
        end_minutes_txt = driver.find_element_by_id("xmm0")
        if end_minutes_txt.get_attribute("value") == "":
            end_minutes_txt.clear()
            end_minutes_txt.send_keys(end_minutes)
    next_page = driver.find_element_by_xpath("//a[starts-with(@href, 'javascript: do_submit2')]").click()
    page_date = datetime.strptime(driver.find_element_by_name("d").get_attribute("value"), "%Y-%m-%d")
    small_day_txt = driver.find_element_by_xpath("//form/table/tbody/tr[7]/td/table/tbody/tr/td[2]/font[2]/b").text
    absence_reason = driver.find_element_by_xpath("//form/table/tbody/tr[8]//table/tbody/tr[8]//table//tr[1]//select/option[@selected]").get_attribute("value")
    page_month = page_date.month
driver.find_element_by_name("B1").submit()
driver.switch_to.window(driver.window_handles[0])
driver.close()
driver.quit()
