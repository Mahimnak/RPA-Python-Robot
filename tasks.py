from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
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
    browser.configure(
        slowmo=100,
    )
    open_robot_order_site()
    close_annoying_modal()
    download_file()
    archive_receipts()

def open_robot_order_site():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('Ok')")

def download_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    for row in orders:
        fill_the_form(row)
    
def fill_the_form(details):
    """Fill the form by taking the details from the orders csv file stored in a variable"""
    page = browser.page()
    page.select_option("#head", str(details["Head"]))
    if str(details["Body"])=='1':
        page.click("#id-body-1")
    elif str(details["Body"])=='2':
        page.click("#id-body-2")
    elif str(details["Body"])=='3':
        page.click("#id-body-3")
    elif str(details["Body"])=='4':
        page.click("#id-body-4")
    elif str(details["Body"])=='5':
        page.click("#id-body-5")
    elif str(details["Body"])=='6':
        page.click("#id-body-6") 
    page.fill(".form-control", str(details["Legs"]))      
    page.fill("#address", str(details["Address"]))
    page.click("#preview")  
    page.click("#order")  
    element = page.locator('.alert').count()
    while element==1:
        page.click("#order")
        element = page.locator('.alert')
    store_receipt_as_pdf(details["Order number"])
    screenshot_robot(details["Order number"])
    path1 = "O:/RPA/Robot2/output/receipts/"+details["Order number"]+".pdf"
    path2 = "O:/RPA/Robot2/output/receipts/"+details["Order number"]+".png"
    embed_screenshot_to_receipt(path2,path1)
    page.click("#order-another")  
    close_annoying_modal()

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    neworder = page.locator("#receipt").inner_html()

    pdf = PDF() 
    pdf.html_to_pdf(neworder, "O:/RPA/Robot2/output/receipts/"+order_number+".pdf")
    path = "O:/RPA/Robot2/output/receipts/"+order_number+".pdf"

def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path="O:/RPA/Robot2/output/receipts/"+order_number+".png")
    path = "O:/RPA/Robot2/output/receipts/"+order_number+".png"

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    list_of_files = [
        screenshot
    ]
    pdf.add_files_to_pdf(
        files = list_of_files,
        target_document = pdf_file
    )

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip('O:/RPA/Robot2/output/receipts', 'receipts.zip', exclude ='*.png')