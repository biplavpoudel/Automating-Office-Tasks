from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=1000,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_and_submit_sales_form()

def open_the_intranet_website():
    """Navigates to the given url"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Logs into Admin account by entering credentials"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sales_form():
    """Fills in sales data and submits"""
    page = browser.page()
    page.fill("#firstname","Biplav")
    page.fill("#lastname","Poudel")
    page.select_option("#salestarget","10000")
    page.fill("#salesresult","123")
    page.click("button:text('Submit')")

def download_excel_file():
    """Download Excel file from the url"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
