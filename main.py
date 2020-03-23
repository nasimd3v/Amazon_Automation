import os
import sys
import time
from datetime import datetime
from os import path

import urllib3
import xlsxwriter
from selenium.common.exceptions import NoSuchWindowException, WebDriverException
from xlrd import open_workbook


header = {
    "A1": "URL",
    "B1": "Product Name",
    "C1": "Price",
    "D1": "Brand",
    "E1": "Sold by",
    "F1": "Fulfilled by",
    "G1": "Date",
    "H1": "Time",
    "I1": "QTY",
    "J1": "Rating",
    "K1": "Number of Reviews",
    "L1": "BSR",
    "M1": "Category",
    "N1": "Number of OS",
    "O1": "Other SP"
}


def result_file_creator(file_name):
    if not path.exists("result_file"):
        os.mkdir("result_file")
    files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    final_result_file = str("result_file/result-of-{}").format(file_name)
    if final_result_file not in files:
        # creating result file
        workbook = xlsxwriter.Workbook(final_result_file)
        worksheet = workbook.add_worksheet("result")

        for pos, name in header.items():
            worksheet.write(
                pos, name
            )
        workbook.close()


def read_test_file(file_name):
    print(f"File Scanning Start : {file_name}")
    from engine import engine_, printer, browser
    printer([list(header.values())])
    wb = open_workbook(file_name)
    for s in wb.sheets():
        for row in range(0, s.nrows):
            col_names = s.row(0)
            for name, col in zip(col_names, range(s.ncols)):
                url = s.cell(row, col).value
                # try:
                try:
                    engine_(url, file_name, row + 1)
                except urllib3.exceptions.ProtocolError:
                    print("Internet Connection Problem")
                except NoSuchWindowException:
                    print("Chrome Widow Close!")
                    sys.exit()
                except KeyboardInterrupt:
                    print("Program Terminate!")
                except WebDriverException:
                    print("Chrome Not Reachable!")
                except Exception as E:
                    print(E)
    # browser.quit()
    print("__________________________________________________________________")
    print("All URL Scraped. Result Save in (result-of-{})".format(file_name))
    print("Waiting For A New File")


def task(file_name):

    now = datetime.now()
    Time = int(now.strftime("%H"))
    Date = int(datetime.today().strftime('%m%d'))

    if not path.exists("task"):
        os.mkdir("task")
    # checking if task file exists

    t_file = str(file_name).replace(".xlsx", ".txt")
    if not path.isfile("task/" + t_file):
        f = open("task/" + t_file, "w")
        f.write(str(Time) + " " + str(Date) + " " + str(1))
        f.close()

        result_file_creator(file_name)
        read_test_file(file_name)

    else:
        f = open("task/" + t_file, "r+")
        data = [int(x) for x in f.readline().split(" ")]
        if data[0] <= Time and data[1] < Date and data[2] < 4:
            Time = int(now.strftime("%H"))
            Date = int(datetime.today().strftime('%m%d'))
            result_file_creator(file_name)
            read_test_file(file_name)
            f.seek(0)
            f.write(str(Time) + " " + str(Date) + " " + str(data[2] + 1))
            f.close()


if __name__ == '__main__':
    print("Waiting For a File:")
    while True:
        time.sleep(0.05)
        files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        if len(files) != 0:
            for file in files:
                task(file)
        time.sleep(0.05)
