from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

from playwright.sync_api import TimeoutError

from PIL import Image

import os

@task
def get_robot_order_list():
    """
    Gets csv with robot order list.
    """
    # browser.configure(
    #     slowmo=500
    # )
    navigate_to_robocorp_site()
    get_orders_file()


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSPareBin Industries Inc.
    Saves the order HTML receipts as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    # browser.configure(
    #     slowmo=500
    # )
    setup_environment()
    navigate_to_robotsparebin_site()
    enter_all_orders()
    archive_order_receipts()


def setup_environment():
    """Create all necessary folders"""
    os.mkdir("output/screenshots/")
    os.mkdir("output/html_receipts/")
    os.mkdir("output/complete_receipts/")


def navigate_to_robocorp_site():
    """Navigate to the given URL."""
    browser.goto("https://robocorp.com/docs/courses/build-a-robot-python/rules-for-the-robot")


def navigate_to_robotsparebin_site():
    """Navigate to the given URL."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    browser.page().click(selector="//button[@class='btn btn-dark']")


def order_another_robot():
    """Click Order Another Robot in order process"""
    page = browser.page()

    page.click(selector="#order-another")
    page.click(selector="//button[@class='btn btn-dark']")


def get_orders_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def enter_all_orders():
    tables = Tables()
    page = browser.page()

    orders = tables.read_table_from_csv("orders.csv", header=True)

    for order in orders:
        order_number = order["Order number"]
        head = str(int(order["Head"]) + 1)
        body = order["Body"]
        legs = order["Legs"]
        address = order["Address"]

        enter_single_order_info(head=head, body=body, legs=legs, address=address)

        page.click(selector="#order")

        save_html_receipt_as_pdf(order_number=order_number)
        get_robot_images(order_number=order_number)
        combine_robot_images(order_number=order_number)

        embed_screenshot_into_pdf(order_number=order_number)

        order_another_robot()


def enter_single_order_info(head, body, legs, address):
    """Fill out form with order info"""
    page = browser.page()

    head_option_element = page.locator(f"#head > option:nth-child({head})")
    head_text = head_option_element.text_content()
    page.select_option("#head", label=head_text)

    page.click(f"input[name='body'][value='{body}']")

    page.type("//*[@class='form-control']", legs)

    page.type("#address", address)


def save_html_receipt_as_pdf(order_number):
    """Save the HTML receipt of a robot order as a PDF"""
    page = browser.page()
    page.set_default_timeout(3000)

    order_error_present = True

    while order_error_present:
        try:
            robot_order_html = page.locator(selector="#receipt").inner_html()
            order_error_present = False
        except TimeoutError:
            print(f"Kon de html ontvangst voor order {order_number} niet binnen de timeout opvragen. Opnieuw proberen...")
            page.click(selector="#order")

    pdf = PDF()
    pdf.html_to_pdf(robot_order_html, f"output/html_receipts/html_receipt_order_{order_number}.pdf")

    page.set_default_timeout(30000)


def get_robot_images(order_number):
    """Save image of ordered robot to add to the receipt."""
    page = browser.page()

    robot_head_element = page.locator("#robot-preview-image > img:nth-child(1)")
    robot_body_element = page.locator("#robot-preview-image > img:nth-child(2)")
    robot_legs_element = page.locator("#robot-preview-image > img:nth-child(3)")

    screenshot_head = browser.screenshot(element=robot_head_element)
    screenshot_body = browser.screenshot(element=robot_body_element)
    screenshot_legs = browser.screenshot(element=robot_legs_element)

    with open(f"output/screenshots/robot_head_{order_number}.png", "wb") as f:
        f.write(screenshot_head)
    with open(f"output/screenshots/robot_body_{order_number}.png", "wb") as f:
        f.write(screenshot_body)
    with open(f"output/screenshots/robot_legs_{order_number}.png", "wb") as f:
        f.write(screenshot_legs)


def embed_screenshot_into_pdf(order_number):
    """Concatenates robot images and embeds them to the receipt pdf"""
    pdf = PDF()

    pdf.add_files_to_pdf(files=[f"output/html_receipts/html_receipt_order_{order_number}.pdf", f"output/screenshots/combined_robot_{order_number}.png"],
                         target_document=f"output/complete_receipts/robot_order_receipt_{order_number}.pdf")


def combine_robot_images(order_number):
    # Open images
    head_image = Image.open(f"output/screenshots/robot_head_{order_number}.png")
    body_image = Image.open(f"output/screenshots/robot_body_{order_number}.png")
    legs_image = Image.open(f"output/screenshots/robot_legs_{order_number}.png")

    # Calculate the total height and the maximum width from all images
    total_height = head_image.height + body_image.height + legs_image.height
    max_width = max(head_image.width, body_image.width, legs_image.width)

    # Create a new blank image with the calculated total height and max width
    combined_image = Image.new('RGBA', (max_width, total_height))

    # Paste images one below the other
    current_height = 0
    for image in [head_image, body_image, legs_image]:
        combined_image.paste(image, (0, current_height))
        current_height += image.height

    combined_image_path = f"output/screenshots/combined_robot_{order_number}.png"
    combined_image.save(combined_image_path)


def archive_order_receipts():
    lib = Archive()
    lib.archive_folder_with_zip(folder="output/complete_receipts/", archive_name="output/robot_order_receipts.zip")
