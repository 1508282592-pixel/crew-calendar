import os
import base64
from playwright.sync_api import sync_playwright
from PIL import Image
import pytesseract
from ics import Calendar, Event
from datetime import datetime, timedelta

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]


def solve_captcha(page):

    img = page.locator("img").nth(0).get_attribute("src")

    if img and img.startswith("data:image"):
        base64_data = img.split(",")[1]

        img_bytes = base64.b64decode(base64_data)

        with open("captcha.jpg", "wb") as f:
            f.write(img_bytes)

        image = Image.open("captcha.jpg")

        code = pytesseract.image_to_string(image)

        code = code.strip().replace(" ", "")

        print("captcha:", code)

        return code

    return ""


def login(page):

    page.goto("https://cp.9cair.com")

    page.wait_for_timeout(5000)

    page.fill('input[name="username"]', USERNAME)
    page.fill('input[name="password"]', PASSWORD)

    code = solve_captcha(page)

    page.fill('input[name="validcode"]', code)

    page.click("text=Login")

    page.wait_for_timeout(5000)


def create_calendar():

    c = Calendar()

    e = Event()

    e.name = "Crew Schedule Sync"

    e.begin = datetime.now()

    e.end = datetime.now() + timedelta(hours=1)

    e.description = "Auto generated schedule"

    c.events.add(e)

    with open("crew_schedule.ics", "w") as f:

        f.writelines(c)


def run():

    with sync_playwright() as p:

        browser = p.chromium.launch()

        page = browser.new_page()

        login(page)

        create_calendar()

        browser.close()


if __name__ == "__main__":
    run()
