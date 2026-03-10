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
        return code.strip().replace(" ", "")

    return ""


def login(page):
    page.goto("https://cp.9cair.com")
    page.wait_for_timeout(5000)

    inputs = page.locator("input")
    inputs.nth(0).fill(USERNAME)
    inputs.nth(1).fill(PASSWORD)

    code = solve_captcha(page)
    inputs.nth(2).fill(code)

    page.click("text=Login")
    page.wait_for_timeout(5000)


def open_my_tasks(page):
    page.click("text=我的任务")
    page.wait_for_timeout(3000)

    # 点击所有展开箭头
    arrows = page.locator("svg")  # 小箭头通常是svg图标
    count = arrows.count()

    for i in range(count):
        try:
            arrows.nth(i).click()
        except:
            pass

    page.wait_for_timeout(2000)


def create_calendar_from_page(page):

    content = page.content()

    c = Calendar()

    e = Event()
    e.name = "Crew Schedule Loaded"
    e.begin = datetime.now()
    e.end = datetime.now() + timedelta(hours=1)
    e.description = "Successfully loaded My Tasks page"

    c.events.add(e)

    with open("crew_schedule.ics", "w") as f:
        f.writelines(c)


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        login(page)
        open_my_tasks(page)
        create_calendar_from_page(page)

        browser.close()


if __name__ == "__main__":
    run()
