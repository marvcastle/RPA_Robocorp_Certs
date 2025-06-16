from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os
import time
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
        slowmo=1000,
    )
    open_website()
    pass_pop_up()
    #web_download_csv()
    orders = read_csv_file()
    process_order(orders)
    

def open_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def pass_pop_up():
    """Closes the pop-up window if it appears"""
    try:
        page = browser.page()
        page.click("button:text('I guess so...')")
    except browser.ElementNotFound:
        pass  # Pop-up not found, continue
    # You can target HTML elements either with CSS type, class, ID, or attribute selectors or XPath expressions, or a combination of CSS, XPath, and text selectors.

def web_download_csv():
    """Downloads the csv file from the website"""
    http = HTTP()
    http.download(
        "https://robotsparebinindustries.com/orders.csv",
        overwrite=True
    )

def read_csv_file():
    """Reads the downloaded CSV file and returns its content"""
    # files = Files()
    # files.open_workbook("orders.csv")
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv", header=True)
    return orders

def get_model_mapping():
    """Extracts the model name to part number mapping from the web page table."""
    page = browser.page()
    page.click("button:text('Show model info')")
    rows = page.locator("#model-info tr").all()
    mapping = {}
    for row in rows[1:]:  # Skip header row
        cells = row.locator("td").all_inner_texts()
        if len(cells) == 2:
            model_name, part_number = cells
            mapping[part_number] = model_name
    return mapping

def fill_form(order):
    page = browser.page()
    page.select_option("#head", order["Head"])
    page.click(f"input[name='body'][value='{order['Body']}']")
    page.fill("input[placeholder*='legs']", str(order["Legs"]))
    page.fill("#address", order["Address"])

def submit_order():
    page = browser.page()
    page.click("#preview")
    page.click("#order")

def get_order_number():
    page = browser.page()
    order_number = page.locator("#receipt .badge-success").inner_text()
    return order_number

def process_order(orders):
    page = browser.page()
    for order in orders:
        for attempt in range(5):  # Try twice before giving up
            try:
                fill_form(order)
                submit_order()
                # Wait for either success or error
                if page.locator("div.alert-danger").is_visible(timeout=500):
                    print("Order error detected, retrying...")
                    page.click("#order")
                    continue  # Retry the same order
                else:
                    # Success: click "Order another robot" to reset form
                    order_number = get_order_number()
                    document_order(order_number)
                    page.click("#order-another")
                    pass_pop_up()
                    break  # Move to next order
            except Exception as e:
                print(f"Unexpected error: {e}")
                page.reload()
                pass_pop_up()
                continue

def error_handling():
    """Handles errors that may occur during the task execution"""

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_file = f"output/receipt_{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_file)
    return pdf_file

def screenshot_robot(order_number):
    page = browser.page()
    screenshot_file = f"output/robot_{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_file)
    return screenshot_file

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )

def archive_receipts(order_number):
    """Creates a ZIP archive of the receipts and images"""
    archive = Archive()
    files = [
        f"receipt_{order_number}.pdf",
        f"robot_{order_number}.png"
    ]
    archive.archive_folder_with_zip(
        folder='./output',        
        archive_name=f'output/robot_receipt_{order_number}.zip', 
        include=f'receipt_{order_number}.pdf'
        # f"output/receipt_{order_number}.pdf",
        # f"output/robot_{order_number}.png"
    )

def add_to_zip(order_number):
    """Creates a ZIP archive of the receipts and images"""
    archive = Archive()
    files = [
        # f"output/receipt_{order_number}.pdf",
        f"output/robot_{order_number}.png"
    ]
    archive.add_to_archive(
        files=f"output/robot_{order_number}.png",
        archive_name=f"output/robot_receipt_{order_number}.zip"
     )

def document_order(order_number):
    pdf_file = store_receipt_as_pdf(order_number)
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot, pdf_file)
    wait_for_file(pdf_file)
    wait_for_file(screenshot)
    archive_receipts(order_number)
    add_to_zip(order_number)

def wait_for_file(filepath, timeout=5):
    """Wait up to `timeout` seconds for a file to exist."""
    start = time.time()
    while not os.path.exists(filepath):
        if time.time() - start > timeout:
            raise TimeoutError(f"File {filepath} not found after {timeout} seconds.")
        time.sleep(0.1)

        # head_item_id = order["Head"]
        # head_model_name = model_mapping.get(head_item_id, "")
        # # Now you can use head_model_name or head_item_id as needed
        # # Example: select the head by part number
        # page.select_option("#head", head_model_name)
        # # Continue filling the rest of the form...
