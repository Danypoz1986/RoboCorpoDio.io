from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from bs4 import BeautifulSoup
import re
import time

# Open the browser and navigate to the robot order page
def open_robot_order_website():
    browser = Selenium()
    browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")
    return browser

# Download the orders CSV file
def get_orders():
    http = HTTP()
    tables = Tables()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    orders = tables.read_table_from_csv("orders.csv")
    return orders

# Close the annoying modal by clicking on one of the buttons
def close_annoying_modal(browser):
    try:
        # Wait for the modal to appear
        browser.wait_until_element_is_visible("xpath://div[contains(@class, 'modal')]", timeout=10)
    except Exception as e:
        print("Modal did not appear: ", e)
        return

    # List of all possible buttons in the modal
    possible_buttons = ["OK", "Yep", "I guess so...", "No way!"]

    # Attempt to click each button until one works
    for button_text in possible_buttons:
        try:
            button_xpath = f"//button[contains(text(), \"{button_text}\")]"
            browser.wait_until_element_is_visible(button_xpath, timeout=3)
            browser.click_element(button_xpath)
            print(f"Clicked modal button: {button_text}")
            time.sleep(2)  # Give some time to ensure the modal is gone
            return
        except Exception as e:
            print(f"Button '{button_text}' not found or not clickable: {e}. Continuing to try other buttons.")
    print("No buttons were clicked on the modal.")
    browser.capture_page_screenshot("output/modal_not_dismissed.png")

# Fill out the form with the given order details
def fill_the_form(browser, order):
    try:
        # Select Head
        browser.scroll_element_into_view("id:head")
        browser.wait_until_element_is_visible("id:head", timeout=5)
        browser.select_from_list_by_value("id:head", order['Head'])

        # Click Body
        body_value = order['Body']
        browser.scroll_element_into_view(f"xpath://input[@value='{body_value}' and @name='body']")
        browser.wait_until_element_is_visible(f"xpath://input[@value='{body_value}' and @name='body']", timeout=5)
        browser.click_element(f"xpath://input[@value='{body_value}' and @name='body']")

        # Input Legs
        browser.scroll_element_into_view("css:input.form-control")
        browser.wait_until_element_is_visible("css:input.form-control", timeout=5)
        browser.input_text("css:input.form-control", order['Legs'])

        # Input Address
        browser.scroll_element_into_view("id:address")
        browser.wait_until_element_is_visible("id:address", timeout=5)
        browser.input_text("id:address", order['Address'])

        print("Form filled successfully")
    except Exception as e:
        print(f"Failed during filling the form: {e}. Taking a screenshot for debugging...")
        browser.capture_page_screenshot("output/fill_the_form_error.png")

# Click on the preview button
def preview_robot(browser):
    browser.scroll_element_into_view("id:preview")
    browser.wait_until_element_is_visible("id:preview", timeout=5)
    browser.click_element("id:preview")

# Submit the order form
def submit_order(browser):
    try:
        retries = 5  # Increase retries to give more attempts
        for attempt in range(retries):
            browser.scroll_element_into_view("id:order")
            browser.wait_until_element_is_visible("id:order", timeout=5)
            try:
                browser.click_element("id:order")
                print(f"Submit button clicked on attempt {attempt + 1}")
                time.sleep(5)  # Wait to see if the submission is successful

                # Check if any known error messages appear, if so, retry
                error_messages = [
                    "External Server Error",
                    "Who Came Up With These Annoying Errors?!",
                    "Server Out Of Ink Error",
                    "Guess what? A server Error!"
                ]

                error_found = False
                for error_message in error_messages:
                    if browser.is_element_visible(f"xpath://div[contains(text(), '{error_message}')]"):
                        print(f"{error_message} encountered, retrying...")
                        error_found = True
                        break

                if error_found:
                    if attempt < retries - 1:
                        browser.reload_page()  # Reload the page to reset state
                        close_annoying_modal(browser)  # Close the modal again
                    continue
                else:
                    print("Order submitted successfully.")
                    break
            except Exception as e:
                print(f"Failed to click submit button on attempt {attempt + 1}: {e}")
                browser.capture_page_screenshot(f"output/submit_order_error_attempt_{attempt + 1}.png")
                time.sleep(3)
                if attempt == retries - 1:
                    raise e
    except Exception as e:
        print(f"Failed to click submit order button after retries: {e}. Taking a screenshot for debugging...")
        browser.capture_page_screenshot("output/submit_order_error_final.png")

# Store the receipt as a PDF
def store_receipt_as_pdf(browser, order_number):
    pdf = PDF()
    html_content = browser.get_source()

    # Clean up the HTML using BeautifulSoup and remove images
    soup = BeautifulSoup(html_content, "html.parser")
    for img in soup.find_all("img"):
        img.decompose()  # Remove all <img> tags completely

    # Convert the cleaned HTML to a string
    cleaned_html = soup.prettify()

    # Remove emojis and other non-ASCII characters using regex
    cleaned_html = re.sub(r'[^\x00-\x7F]+', '', cleaned_html)
    
    # Save the cleaned HTML to a file for inspection
    with open(f"output/cleaned_html_{order_number}.html", "w", encoding="utf-8") as file:
        file.write(cleaned_html)

    # Create the PDF with cleaned HTML
    receipt_path = f"output/receipt_{order_number}.pdf"
    pdf.html_to_pdf(cleaned_html, receipt_path)
    return receipt_path

# Take a screenshot of the robot
def screenshot_robot(browser, order_number):
    screenshot_path = f"output/screenshot_{order_number}.png"
    browser.capture_page_screenshot(screenshot_path)
    return screenshot_path

# Embed the screenshot into the receipt PDF
def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf([pdf_file, screenshot], pdf_file)

# Click the 'Order Another Robot' button
def order_another_robot(browser):
    try:
        browser.scroll_element_into_view("id:order-another")
        browser.wait_until_element_is_visible("id:order-another", timeout=15)
        retries = 3
        for attempt in range(retries):
            try:
                browser.click_element("id:order-another")
                print(f"Successfully clicked 'Order Another Robot' button on attempt {attempt + 1}")
                time.sleep(2)
                break
            except Exception as e:
                print(f"Failed to click 'Order Another Robot' button on attempt {attempt + 1}: {e}")
                browser.capture_page_screenshot(f"output/order_another_robot_error_attempt_{attempt + 1}.png")
                time.sleep(2)
                if attempt == retries - 1:
                    raise e
    except Exception as e:
        print(f"Failed to click 'Order Another Robot' button after multiple attempts: {e}. Taking a screenshot for debugging...")
        browser.capture_page_screenshot("output/order_another_robot_not_found.png")

# Archive all receipts into a ZIP file
def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip("output", "output/receipts.zip")

@task
def main():
    # Step 1: Open the robot order website
    browser = open_robot_order_website()

    # Step 2: Download the orders file
    orders = get_orders()

    # Step 3: Process each order
    for order in orders:
        close_annoying_modal(browser)
        fill_the_form(browser, order)
        preview_robot(browser)
        submit_order(browser)
        
        # Step 4: Save receipt and screenshot
        order_number = order['Order number']
        receipt_path = store_receipt_as_pdf(browser, order_number)
        screenshot_path = screenshot_robot(browser, order_number)
        
        # Step 5: Embed screenshot into receipt
        embed_screenshot_to_receipt(screenshot_path, receipt_path)
        
        # Step 6: Order another robot
        order_another_robot(browser)

    # Step 7: Create a ZIP archive of all receipts
    archive_receipts()

    # Step 8: Close the browser
    browser.close_browser()

if __name__ == "__main__":
    main()
