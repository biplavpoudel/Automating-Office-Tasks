from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=10,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()

def open_the_intranet_website():
    """Navigates to the given url"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Logs into Admin account by entering credentials"""
    page = browser.page()
    page.fill('//*[@id="username"]', "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_sales_form(sales_rep):
    """Fills in sales data and submits"""
    page = browser.page()
    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("button:text('Submit')")

def download_excel_file():
    """Download Excel file from the url"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    """Read data from Excel file and fill in the sales form"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Screenshot the page"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def log_out():
    """Presses log out button"""
    page = browser.page()
    page.click("button:text('Log out')")

def export_as_pdf():
    """Creates a pdf file from sales table"""
    page = browser.page()
    sales_result_html = page.locator('//*[@id="sales-results"]').inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_result_html, "output/sales_results.pdf")