from csv import DictReader

from RPA.Archive import Archive
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables, Table
from robocorp.tasks import task
from selenium.common import ElementClickInterceptedException

browser = Selenium()


@task
def order_robots():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    out_path = "output/receipts"
    open_order_site()
    order_data = download_and_read_csv_data()
    for i, row in enumerate(order_data):
        browser.click_button("//*[contains(text(), 'OK')]")
        fill_and_submit_order_form(row)
        filename_png = take_screenshot(i, out_path)
        filename_pdf = export_order_as_pdf(i, out_path)
        embed_screenshot_to_receipt(out_path, filename_png, filename_pdf)

        if i < len(order_data):
            browser.click_button("//*[contains(text(), 'Order another')]")

    archive_receipts()
    browser.close_browser()


def open_order_site():
    """Opens the intranet page."""
    browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")


def fill_and_submit_order_form(order_data: dict[int | str]):
    """Fill the order form."""
    browser.select_from_list_by_index("id:head", order_data["Head"])
    browser.select_radio_button("body", order_data["Body"])
    browser.input_text("css:input.form-control", order_data["Legs"])
    browser.input_text("id:address", order_data["Address"])
    while not browser.does_page_contain_button("id:order-another"):
        if browser.does_page_contain_button("id:order"):
            try:
                browser.click_button("//*[contains(text(), 'Order')]")
                #browser.click_button("//*[contains(text(), 'Order')]")
            except ElementClickInterceptedException:
                browser.wait_and_click_button("id:order")


def download_and_read_csv_data() -> Table:
    """
    Get the order data from a CSV file but do not save it to disk.

    :return: Order data as a Table.
    """
    http = HTTP()
    res = http.download("https://robotsparebinindustries.com/orders.csv")
    csv_reader = DictReader(res.text.split("\n"), delimiter=",")
    csv = list(csv_reader)
    table = Tables()
    return table.create_table(csv)


# async def wait_until_images_are_loaded(image: WebElement):
#     """Wait until the images are loaded."""
#     image_id = "image-" + image.get_attribute("src").split("/")[-2]
#     browser.assign_id_to_element(image, image_id)
#     browser.wait_for_condition(
#         f"""
#         return document.getElementById("{image_id}").complete &&
#             typeof document.getElementById("{image_id}").naturalWidth != "undefined" &&
#             document.getElementById("{image_id}").naturalWidth > 0;
#         """
#     )


def take_screenshot(order_no: int, out_path: str) -> str:
    """
    Take a screenshot of the ordered robot.

    :return: The Filename of the screenshot.
    """
    filename = f"order-{order_no}.png"
    images = browser.get_webelements("css:div#robot-preview-image img")
    # TODO: Execute loop in parallel
    # corouines = [wait_until_images_are_loaded(image) for image in images]
    # await asyncio.gather(*corouines)
    for image in images:
        image_id = "image-" + image.get_attribute("src").split("/")[-2]
        browser.assign_id_to_element(image, image_id)
        browser.wait_for_condition(
            f"""
            return document.getElementById("{image_id}").complete &&
                typeof document.getElementById("{image_id}").naturalWidth != "undefined" &&
                document.getElementById("{image_id}").naturalWidth > 0;
            """
        )
    browser.screenshot("id:robot-preview-image", f"{out_path}/{filename}")
    return filename


def export_order_as_pdf(order_no: int, out_path: str) -> str:
    """
    Export the order as PDF.

    :return: The Filename of the PDF.
    """
    sales_results_html = browser.get_element_attribute("id:receipt", "outerHTML")
    filename = f"order-{order_no}.pdf"
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, f"{out_path}/{filename}")
    return filename


def embed_screenshot_to_receipt(out_path: str, png_name: str, pdf_name: str):
    """
    Embed the screenshot to the PDF receipt.
    """
    png_path = f"{out_path}/{png_name}"
    pdf_path = f"{out_path}/{pdf_name}"

    list_of_files = [f"{png_path}:align=center"]

    pdf = PDF()
    pdf.add_files_to_pdf(files=list_of_files, target_document=pdf_path, append=True)


def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip(
        folder="output/receipts", archive_name="output/receipts.zip", exclude="*.png"
    )
