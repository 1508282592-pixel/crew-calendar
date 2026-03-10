import os
import re
import io
import base64
from datetime import datetime, timedelta

from PIL import Image, ImageOps
import pytesseract
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from ics import Calendar, Event

LOGIN_URL = "https://cp.9cair.com"
MISSION_URL = "https://cp.9cair.com/html/task/mission.html"

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]


def clean_code(text: str) -> str:
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text[:4]


def preprocess_captcha(img_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(img_bytes)).convert("L")
    img = ImageOps.autocontrast(img)
    img = img.point(lambda x: 255 if x > 165 else 0, mode="1")
    img = img.resize((img.width * 3, img.height * 3))
    return img


def extract_captcha_bytes(page) -> bytes:
    # 优先找带 base64 的验证码图
    imgs = page.locator("img")
    count = imgs.count()

    for i in range(count):
        try:
            src = imgs.nth(i).get_attribute("src", timeout=1000)
            if src and src.startswith("data:image"):
                base64_data = src.split(",", 1)[1]
                return base64.b64decode(base64_data)
        except:
            pass

    raise RuntimeError("未找到验证码图片")


def solve_captcha(page) -> str:
    img_bytes = extract_captcha_bytes(page)
    processed = preprocess_captcha(img_bytes)

    processed.save("captcha_processed.png")

    config = r'--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    raw = pytesseract.image_to_string(processed, config=config)
    code = clean_code(raw)

    print("captcha raw:", raw)
    print("captcha cleaned:", code)

    return code


def fill_login_form(page, code: str):
    inputs = page.locator("input")
    count = inputs.count()
    if count < 3:
        raise RuntimeError(f"登录页输入框数量异常: {count}")

    inputs.nth(0).fill(USERNAME)
    inputs.nth(1).fill(PASSWORD)
    inputs.nth(2).fill(code)


def login(page, max_retries: int = 5):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(4000)

    for attempt in range(1, max_retries + 1):
        try:
            code = solve_captcha(page)

            if len(code) != 4:
                print(f"第 {attempt} 次验证码长度异常: {code}")
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(2500)
                continue

            fill_login_form(page, code)
            page.click("text=Login")
            page.wait_for_timeout(5000)

            current_url = page.url
            body_text = page.locator("body").inner_text(timeout=5000)

            # 登录成功：不再停留在登录页，或者正文出现机组门户相关字样
            if ("统一认证中心" not in body_text) and ("Login" not in body_text):
                print(f"登录成功，attempt={attempt}, url={current_url}")
                return

            print(f"第 {attempt} 次登录疑似失败，准备重试")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)

        except Exception as e:
            print(f"第 {attempt} 次登录异常: {e}")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)

    raise RuntimeError("多次尝试后仍无法登录")


def open_mission_page(page):
    page.goto(MISSION_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(6000)


def expand_all(page):
    selectors = [
        "svg",
        "i",
        "button",
        "[class*='arrow']",
        "[class*='expand']",
        "[class*='toggle']",
    ]

    for selector in selectors:
        try:
            items = page.locator(selector)
            count = items.count()
            for i in range(count):
                try:
                    items.nth(i).click(timeout=800)
                    page.wait_for_timeout(250)
                except:
                    pass
        except:
            pass

    page.wait_for_timeout(2000)


def create_calendar_from_page(page):
    text = page.locator("body").inner_text()

    c = Calendar()

    e = Event()
    e.name = "Crew Tasks Loaded"
    e.begin = datetime.now()
    e.end = datetime.now() + timedelta(hours=1)
    e.description = text[:3000]

    c.events.add(e)

    with open("crew_schedule.ics", "w", encoding="utf-8") as f:
        f.writelines(c)


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})

        login(page, max_retries=5)
        open_mission_page(page)
        expand_all(page)
        create_calendar_from_page(page)

        browser.close()


if __name__ == "__main__":
    run()
