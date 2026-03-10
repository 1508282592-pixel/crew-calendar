import os
import base64
import requests
from playwright.sync_api import sync_playwright
from PIL import Image
import pytesseract
from ics import Calendar, Event
from datetime import datetime, timedelta

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]

def solve_captcha(page):
    # 获取验证码图片
    img = page.locator("img").nth(0).get_attribute("src")

    if img.startswith("data:image"):
        base64_data = img.split(",")[1]
        img_bytes = base64.b64decode(base64_data)

        with open("captcha.jpg", "wb") as f:
            f.write(img_bytes)

        image = Image.open("captcha.jpg")
        code = pytesseract.image_to_string(image).strip()

        return code
    return ""

def login(page):
    page.goto("https://cp.9cair.com")

    page.fill('input[placeholder="Username"]', USERNAME)
    page.fill('input[placeholder="Password"]', PASSWORD)

    captcha = solve_captcha(page)
    page.fill('input[placeholder="Validcode"]', captcha)

    page.click("text=Login")
    page.wait_for_timeout(5000)

def create_test_calendar():
    c = Calendar()
    e = Event()
    e.name = "Crew Sync Test"
    e.begin = datetime.now()
    e.end = datetime.now() + timedelta(hours=1)
    c.events.add(e)

    with open("crew_schedule.ics", "w") as f:
        f.writelines(c)

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        login(page)

        create_test_calendar()

        browser.close()

if __name__ == "__main__":
    run()
