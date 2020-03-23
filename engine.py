import re
import sys
import time
import selenium
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from openpyxl import load_workbook
from texttable import Texttable

exc = "Exception"


def printer(data):
    table = Texttable()
    load_data = table.add_rows(data)
    load_data.set_cols_width([
        4, 15, 8, 8,
        9, 8, 9,
        6, 6, 6,
        8, 6, 15,
        7, 10
    ]
    )
    t_style = load_data.draw()
    print(t_style)


def write_xl_file(data, final_result_file):
    result_file = "result_file/result-of-" + final_result_file
    try:
        open_xl_file = load_workbook(str(result_file))
        worksheet = open_xl_file.worksheets[0]
        worksheet.append(data)
        open_xl_file.save(result_file)
    except FileNotFoundError as e:
        print(" Couldn't Save ! Result File Couldn't Find.")
    except PermissionError:
        print("Result File Open In Another Process ! Please Close Before Continue")
        input("Press Any Key To Continue.........................................")
        write_xl_file(data, final_result_file)


try:
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    browser = webdriver.Chrome(executable_path=r"chromedriver.exe", options=options)
except selenium.common.exceptions.WebDriverException as w:
    print("(chromedriver) File Not Valid ")
    sys.exit()


def engine_(url, file, counter):
    global Sold_by, QTY, go_cart
    browser.get(url)

    if str("Page Not Found") not in str(browser.title):
        # Gating Product Title
        try:
            Product_Name = str(browser.find_element_by_id("productTitle").text)
        except (TimeoutException, NoSuchElementException) as e:
            Product_Name = exc
        # Gating Product Price
        try:
            Price = str(browser.find_element_by_id("price_inside_buybox").text)
        except selenium.common.exceptions.NoSuchElementException as e:
            try:
                Price = str(browser.find_element_by_xpath("""//*[@id="unqualifiedBuyBox"]/div/div[2]/span""").text)
            except selenium.common.exceptions.NoSuchElementException as E:
                Price = exc
        try:
            Category = str(
                browser.find_element_by_xpath("""//*[@id="wayfinding-breadcrumbs_feature_div"]/ul/li[1]""").text)
        except selenium.common.exceptions.NoSuchElementException as e:
            Category = exc
        try:
            Rating = (str(browser.find_element_by_id("acrPopover").get_property("title")).split(" "))[0]
        except selenium.common.exceptions.NoSuchElementException as e:
            Rating = exc
        try:
            Number_of_reviews = (str(browser.find_element_by_id("acrCustomerReviewText").text).split(" "))[0]
        except selenium.common.exceptions.NoSuchElementException as e:
            Number_of_reviews = exc

        try:
            BSR = ((str(browser.find_element_by_id("SalesRank").text).split(" "))[3]).replace("#", "")
        except selenium.common.exceptions.NoSuchElementException as e:
            BSR = exc
        now = datetime.now()
        Time = str(now.strftime("%H:%M"))
        Date = str(datetime.today().strftime('%d/%m/%Y'))
        try:
            Brand = str(browser.find_element_by_id("bylineInfo").text)
        except selenium.common.exceptions.NoSuchElementException as e:
            Brand = exc

        # sold by and fulfilled by start
        try:
            soldBy = str(browser.find_element_by_id("merchant-info").text)
            if str("Ships from and sold by Amazon.ca") in soldBy:
                Sold_by = "Amazon"
                Fulfilled_by = "Amazon"
            elif str("Ships from and sold by") in soldBy:
                get = str(soldBy).replace("Ships from and sold by", "")
                Sold_by = str(get).replace(".", "")
                Fulfilled_by = Sold_by
            else:
                Fulfilled_by = "Amazon"
                try:
                    Sold_by = str(browser.find_element_by_id("sellerProfileTriggerId").text)
                except selenium.common.exceptions.NoSuchElementException as e:
                    Sold_by = exc
        except selenium.common.exceptions.NoSuchElementException as e:
            Fulfilled_by = "Amazon"
            Sold_by = exc
            # sold by and fulfilled by end

        # soldBy end
        try:
            Number_of_other_sellers = \
                str(browser.find_element_by_xpath("""//span[@class='olp-text']""").text).split('(', 1)[1].split(')')[0]
        except selenium.common.exceptions.NoSuchElementException as e:
            Number_of_other_sellers = exc
        try:
            Other_sellers_price = str(browser.find_element_by_xpath("""//div[@class='olp-text-box']/span[
            @class='a-color-price' and 2]""").text)
        except selenium.common.exceptions.NoSuchElementException as e:
            Other_sellers_price = exc
        # First Page Data End

        status = str(browser.find_element_by_id("availability").text)

        # trying to click add to card button
        if str("In Stock.") in status or str("Usually ships") in status and str("Currently unavailable.") not in status:
            go_cart = True
            try:
                # WebDriverWait(browser, 2).until(
                # EC.presence_of_element_located((By.XPATH, """//*[@id="add-to-cart-button"]"""))).click()
                browser.find_element_by_xpath("""//*[@id="add-to-cart-button"]""").click()
            except:
                pass
            browser.implicitly_wait(3)
            # Click Only Cart Button
            try:
                # WebDriverWait(browser, 2).until(
                #     EC.presence_of_element_located((By.XPATH, )).click()
                browser.find_element_by_xpath("""//*[@id="hlb-view-cart-announce"]""").click()
            except:
                pass

            browser.implicitly_wait(3)

        elif str("Available from these sellers.") in status:
            go_cart = True
            try:
                browser.find_element_by_id("buybox-see-all-buying-choices-announce").click()
            except NoSuchElementException:
                pass
            browser.implicitly_wait(3)
            try:
                # Add To Card
                WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH, """//*[@id="a-autoid-0"]/span"""))).click()
                browser.implicitly_wait(1)
                # Cart Button
                WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH, """//*[@id="hlb-view-cart-announce"]"""))).click()
            except Exception as e:
                try:
                    WebDriverWait(browser, 2).until(
                        EC.presence_of_element_located((By.XPATH, """//span[@id='a-autoid-5']/span[
                        @class='a-button-inner' and 1]/input[@class='a-button-input' and 1]"""))).click()
                    WebDriverWait(browser, 2).until(
                        EC.presence_of_element_located((By.XPATH, """//*[@id="hlb-view-cart-announce"]"""))).click()
                except:
                    pass

        elif str("In Stock.") not in status and str("Currently unavailable.") in status:
            go_cart = False
            QTY = "unavailable"
            Sold_by = QTY
            Price = QTY
            Number_of_reviews = QTY
            Other_sellers_price = QTY

        elif str("left in stock") in status and str("Only") in status:
            QTY = int(re.search(r'\d+', status).group())
            go_cart = False

        else:
            go_cart = False
            QTY = "unavailable"

        if go_cart:
            #     # Set QTY
            try:
                WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.XPATH, """//*[@id="a-autoid-0"]/span"""))).click()
                try:
                    WebDriverWait(browser, 2).until(
                        EC.presence_of_element_located(
                            (By.XPATH, """//*[@id="a-popover-3"]/div/div/ul/li[11]"""))).click()
                except TimeoutException:
                    try:
                        WebDriverWait(browser, 2).until(
                            EC.presence_of_element_located((By.XPATH, """//a[@id='dropdown1_10']"""))).click()
                    except TimeoutException:
                        pass
            except:
                WebDriverWait(browser, 10).until(
                    EC.presence_of_element_located((By.XPATH, """//*[@id="a-autoid-0"]/span"""))).click()

                WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH, """//*[@id="a-popover-3"]/div/div/ul/li[11]"""))).click()

            try:
                input_item = browser.find_element_by_name("quantityBox")
                input_item.send_keys("999")
            except NoSuchElementException:
                pass
            # clicking update button

            WebDriverWait(browser, 2).until(
                EC.presence_of_element_located((By.XPATH, """//*[@id="a-autoid-1-announce"]"""))).click()
            # gating available QTY

            try:
                new_item = browser.find_element_by_xpath(
                    """//div[1 and @class='a-alert-content']/span[1 and @class='a-size-base']""").text
                while True:
                    try:
                        new_item = str(browser.find_element_by_name("quantityBox").get_attribute("value"))
                        if str(new_item) != "999":
                            QTY = str(new_item)
                            break
                        elif str(new_item) == "999":
                            QTY = "999"
                            break
                    except:
                        QTY = exc
            except NoSuchElementException:
                QTY = str(browser.find_element_by_name("quantityBox").get_attribute("value"))
            browser.implicitly_wait(3)
            try:
                WebDriverWait(browser, 20).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    """//span[@class='a-size-small sc-action-delete']/span[@class='a-declarative' and 1]/input[1]"""))).click()
                # browser.find_element_by_xpath().click()
            except:
                WebDriverWait(browser, 20).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    """//span[@class='a-size-small sc-action-delete']/span[@class='a-declarative' and 1]/input[1]"""))).click()
            time.sleep(1)
        data = [
            url,
            Product_Name,
            Price,
            Brand,
            Sold_by,
            Fulfilled_by,
            Date,
            Time,
            QTY,
            Rating,
            Number_of_reviews,
            BSR,
            Category,
            Number_of_other_sellers,
            Other_sellers_price
        ]
        write_xl_file(data, file)
        printer([[
            counter,
            Product_Name[:15],
            Price,
            Brand,
            Sold_by[:15],
            Fulfilled_by[:15],
            Date,
            Time,
            QTY,
            Rating,
            Number_of_reviews,
            BSR,
            Category,
            Number_of_other_sellers,
            Other_sellers_price
        ]])
    else:
        data = [
            url,
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except"
        ]
        write_xl_file(data, file)
        printer([[
            counter,
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except",
            "except"
        ]])
