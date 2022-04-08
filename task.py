from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
import pandas as pd
from RPA.PDF import PDF

pdf = PDF()

r = HTTP()

browser = Selenium(auto_close=False)


def login():
    browser.open_chrome_browser("https://robotsparebinindustries.com/")
    browser.input_text(locator='//*[@id="username"]', text='maria')
    browser.input_text(locator='//*[@id="password"]', text="thoushallnotpass")
    browser.click_button(locator='//*[@type="submit"]')
    browser.wait_until_element_is_visible(locator='//*[@id="logout"]')


def sales_data(firstname,lastname,sales_target,sales):
    browser.input_text(locator='//*[@id="firstname"]',text=firstname)
    browser.input_text(locator='//*[@id="lastname"]', text=lastname)
    browser.click_element(locator='//*[@id="salestarget"]')
    browser.click_element(locator='//*[@value={0}]'.format(sales_target))
    browser.input_text(locator='//*[@id="salesresult"]', text=sales)
    browser.submit_form()

def buy_robot():
    browser.open_chrome_browser('https://robotsparebinindustries.com/#/robot-order')
    browser.wait_until_element_is_visible(locator="//*[@id='root']/div/div[2]/div/div/div/div/div")
    browser.click_element(locator='//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
    browser.click_element(locator='//*[@id="head"]')
    browser.click_element(locator='//*[@id="head"]/option[7]')
    browser.click_button(locator='//*[@id="id-body-6"]')
    browser.input_text(locator='//*[@class="form-control"]', text='2')
    browser.input_text(locator='//*[@id="address"]', text = "Rohini,Mattupara,Nemmara,Palakkad,Kerala")
    browser.click_button(locator='//*[@id="preview"]')
    browser.click_button(locator='//*[@id="order"]')
    
def download_excel():
    r.download('https://robotsparebinindustries.com/SalesData.xlsx', overwrite=True)


def enter_data():
    df = pd.read_excel("C:\\Users\\al3732\\Downloads\\SalesData.xlsx")
    df['Sales'] = df['Sales'].astype(str)
    df['Sales Target'] = df['Sales Target'].astype(str)
    for i in range(len(df)):
        firstname = df['First Name'][i]
        lastname = df['Last Name'][i]
        sales = df['Sales'][i]
        sales_target = df['Sales Target'][i]
        sales_data(firstname,lastname,sales_target,sales)

    browser.screenshot(locator='//*[@id="root"]/div/div/div/div[2]', filename="C:\\Users\\al3732\\Downloads\\sc.png")
    return print('Entry Successfully Completed')


def create_pdf():
    html_elements = browser.get_element_attribute(locator='//*[@id="sales-results"]', attribute='outerHTML')
    pdf.html_to_pdf(html_elements, 'C:\\Users\\al3732\\Downloads\\results.pdf')

def logout():
    browser.click_button(locator="//*[@id='logout']")
    



def sales_testing():
    browser.click_element(locator='//*[@id="salestarget"]')
    browser.click_element(locator='//*[@value="20000"]')

if __name__ == "__main__":
    login()
    # sales()
    # buy_robot()
    # download_excel()
    enter_data()
    # sales_testing()
    # logout()
    create_pdf()

    