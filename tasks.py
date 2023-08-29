from robocorp.tasks import task
import robocorp.browser as browser
import robocorp.http as http
import robocorp.log as log
from playwright.sync_api import expect
from RPA.PDF import PDF
from RPA.Tables import Tables, Table
from RPA.Archive import Archive
from os import path
from tempfile import gettempdir
from shutil import rmtree

# Global Variables
ORDER_LINK_TEXT = "Order your robot!"
ORDER_MODAL_TEXT = "OK"

GLOBAL_RETRY_COUNT = 10
GLOBAL_RETRY_INTERVAL_MS = 100

TEMPDIR = path.abspath(gettempdir())
RECEIPTS_FOLDER = path.join(TEMPDIR, "receipts")
SCREENSHOTS_FOLDER = path.join(TEMPDIR, "screenshots")

pdf = PDF()
tables = Tables()
archive = Archive()


@task
def order_robots_from_RobotSpareBin() -> None:
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    # Set Variables
    HOME_URL = "https://robotsparebinindustries.com/"
    ORDER_URL = f"{HOME_URL}/#/robot-order"
    CSV_FILE = "orders.csv"
    # Configure Browser
    browser.configure(
        slowmo=100,
    )
    expect.set_options(timeout=GLOBAL_RETRY_INTERVAL_MS)
    # Steps
    open_robot_order_website(ORDER_URL)
    orders = get_orders(HOME_URL, CSV_FILE)
    loop_the_orders(orders)
    archive_receipts(RECEIPTS_FOLDER)
    cleanup_temporary_pdf_directory(TEMPDIR)


def open_robot_order_website(order_url: str) -> None:
    """Navigate to RobotSpareBin Industries Inc"""
    browser.goto(order_url)


def close_annoying_modal() -> None:
    """Close Modal that pops-up when opening the orders page"""
    page = browser.page()
    page.get_by_role("button", name=ORDER_MODAL_TEXT).click()


def get_orders(home_url: str, csv_file: str) -> Table:
    """Get orders from a CSV file from the website"""
    csv_url = f"{home_url}/{csv_file}"
    csv_dest = path.join(TEMPDIR, csv_file)
    http.download(url=csv_url, target_file=csv_dest, overwrite=True)

    orders = tables.read_table_from_csv(
        csv_dest,
        header=True,
    )

    return orders


def loop_the_orders(orders: Table) -> None:
    for order in orders:
        log.info(order)
        close_annoying_modal()
        fill_the_form(order)
        preview_the_robot()
        submit_the_order()
        pdf_file = store_receipt_as_pdf(order["Order number"])
        screenshot = screenshot_robot(order["Order number"])
        embed_screenshot_to_receipt(screenshot, pdf_file)
        order_another_robot()


def fill_the_form(order) -> None:
    page = browser.page()
    page.locator("id=head").select_option(value=order["Head"])
    page.locator(f'id=id-body-{order["Body"]}').click()
    page.get_by_placeholder("Enter the part number for the legs").fill(
        str(order["Legs"])
    )
    page.locator("id=address").fill(order["Address"])


def preview_the_robot():
    page = browser.page()
    page.get_by_text("Preview").click()


def submit_the_order() -> None:
    for i in range(GLOBAL_RETRY_COUNT):
        page = browser.page()
        page.get_by_text("Order", exact=True).click()
        try:
            page = browser.page()
            expect(
                page.get_by_role("button", name="Order another robot"),
            ).to_be_visible()
            return
        except:
            pass

    raise RuntimeError(f"Submitting Order failed after {i + 1} attempts")


def store_receipt_as_pdf(order_number: str) -> str:
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    receipt_pdf = path.join(RECEIPTS_FOLDER, f"{order_number}.pdf")
    pdf.html_to_pdf(receipt_html, receipt_pdf)

    return receipt_pdf


def screenshot_robot(order_number: str) -> str:
    screenshot = path.join(SCREENSHOTS_FOLDER, f"{order_number}.png")
    page = browser.page()
    page.locator("id=robot-preview-image").screenshot(path=screenshot)
    return screenshot


def embed_screenshot_to_receipt(screenshot: str, pdf_file: str) -> None:
    pdf.add_files_to_pdf(files=[pdf_file, screenshot], target_document=pdf_file)


def order_another_robot() -> None:
    page = browser.page()
    page.get_by_role("button", name="Order another robot").click()


def archive_receipts(receipts_dir: str) -> None:
    archive.archive_folder_with_zip(
        receipts_dir, archive_name=path.join("output", "PDFs.zip")
    )


def cleanup_temporary_pdf_directory(tempdir: str) -> None:
    rmtree(tempdir, ignore_errors=True)
