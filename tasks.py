from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.Archive import Archive
@task
def order_robots_from_RobotSpareBin():
   """
   Orders robots from RobotSpareBin Industries Inc.
   Saves the order HTML receipt as a PDF file.
   Saves the screenshot of the ordered robot.
   Embeds the screenshot of the robot to the PDF receipt.
   Creates ZIP archive of the receipts and the images.
   """
   browser.configure(slowmo=100)

   open_robot_order_website()
   orders = download_excel_file_and_get_orders()
   return orders

def open_robot_order_website():
   """Navigates to RobotSpareBin Industries"""
   browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file_and_get_orders():
   """Downloads a CSV file from the URL"""
   ae = 0
   http = HTTP()
   http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
   library = Tables()
   orders = library.read_table_from_csv(
    "orders.csv"
)
   for order in orders:
      fill_and_submit_order_form(order)

   return ae

def close_annoying_model():
   """Get rid of that annoying pop up"""
   page = browser.page()
   page.click("button:text('OK')")
   
def fill_and_submit_order_form(order_data):
   """Fill the orders data"""
   page = browser.page()
   close_annoying_model()
   

   page.select_option("#head", order_data["Head"])
   page.check(f"#id-body-" + order_data["Body"])
   page.fill(".form-control", order_data["Legs"])
   page.fill("#address", order_data["Address"])
   page.click("#preview")
   page.click("#order")
   while not page.query_selector("#order-another"):
        page.click("#order")

   store_receipt_as_pdf(order_data["Order number"])

   page.click("#order-another")

def store_receipt_as_pdf(order_number):
   """Take the screenshot of the receipt"""
   page = browser.page()

   order_receipt = page.locator("#receipt").inner_html()
   pdf = PDF()
   pdf_file = f"output/{order_number}.pdf"
   pdf.html_to_pdf(order_receipt, pdf_file)
   screenshot_robot(order_receipt, pdf_file)
   
def screenshot_robot(order_number,pdf_file):
    """Take the screenshot of the robot"""
    page = browser.page()
    robot = page.query_selector("#robot-preview-image")
    screenshot = f"output/{order_number}.png"
    robot.screenshot(path=screenshot)
    embed_screenshot_to_receipt(screenshot, pdf_file)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot to thge receipt PDF""" 
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def archive_receipts():
   """Archieve the receipts for future use"""
   folder = Archive()
   folder.archive_folder_with_zip('output', 'output/orders.zip', include='*.pdf')




   
         
        
   


    