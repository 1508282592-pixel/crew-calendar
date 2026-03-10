import os
import re
import io
import base64
from itertools import product
from datetime import datetime, timedelta

from PIL import Image, ImageOps, ImageFilter
import pytesseract
from playwright.sync_api import sync_playwright
from ics import Calendar, Event

LOGIN_URL = "https://cp.9cair.com"
MISSION_URL = "https://cp.9cair.com/html/task/mission.html"

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]


def normalize_candidate(text: str) -> str:
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]", "", text)
    if len(text) == 5:
        text = text[:4]
    return text


def score_candidate(text: str) -> int:
    if not text:
        return 0
    score = 0
    if len(text) == 4:
        score += 100
    elif len(text) == 5:
        score += 60
    elif len(text) == 3:
        score += 40
    else:
        score += 10
    score += sum(ch.isalnum() for ch in text)
    return score


def expand_char_options(ch: str):
    mapping = {
        "0": ["0", "O"],
        "O": ["O", "0"],
        "1": ["1", "I", "L"],
        "I": ["I", "1", "L"],
        "L": ["L", "1", "I"],
        "5": ["5", "S"],
        "S": ["S", "5"],
        "8": ["8", "B"],
        "B": ["B", "8"],
        "2": ["2", "Z"],
        "Z": ["Z", "2"],
        "6": ["6", "G"],
        "G": ["G", "6"],
        "3": ["3", "B"],
    }
    return mapping.get(ch, [ch])


def generate_code_candidates(code: str, limit: int = 12):
    pools = [expand_char_options(ch) for ch in code]
    all_codes = []
    for combo in product(*pools):
        cand = "".join(combo)
        if cand not in all_codes:
            all_codes.append(cand)
        if len(all_codes) >= limit:
            break
    return all_codes


def extract_captcha_bytes(page) -> bytes:
    imgs = page.locator("img")
    count = imgs.count()

    for i in range(count):
        try:
            src = imgs.nth(i).get_attribute("src", timeout=1000)
            if src and src.startswith("data:image"):
                return base64.b64decode(src.split(",", 1)[1])
        except:
            pass

    raise RuntimeError("未找到验证码图片")


def build_variants(img_bytes: bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("L")
    img = ImageOps.autocontrast(img)

    variants = []

    variants.append(img.resize((img.width * 3, img.height * 3)))

    for threshold in [140, 155, 170, 185]:
        bw = img.point(lambda x: 255 if x > threshold else 0, mode="1")
        bw = bw.resize((bw.width * 3, bw.height * 3))
        variants.append(bw)

    inv = ImageOps.invert(img)
    inv = inv.resize((inv.width * 3, inv.height * 3))
    variants.append(inv)

    sharp = img.filter(ImageFilter.SHARPEN)
    sharp = sharp.resize((sharp.width * 3, sharp.height * 3))
    variants.append(sharp)

    return variants


def solve_captcha(page) -> str:
    img_bytes = extract_captcha_bytes(page)
    variants = build_variants(img_bytes)

    configs = [
        r'--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        r'--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
        r'--psm 13 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
    ]

    candidates = []

    for idx, variant in enumerate(variants):
        try:
            variant.save(f"captcha_variant_{idx}.png")
        except:
            pass

        for cfg in configs:
            raw = pytesseract.image_to_string(variant, config=cfg)
            cleaned = normalize_candidate(raw)

            print("captcha raw:", repr(raw))
            print("captcha cleaned:", cleaned)

            if cleaned:
                candidates.append(cleaned)

    if not candidates:
        return ""

    candidates = sorted(candidates, key=score_candidate, reverse=True)
    best = candidates[0][:4]

    print("captcha best:", best)
    return best


def fill_login_form(page, code: str):
    inputs = page.locator("input")
    if inputs.count() < 3:
        raise RuntimeError("登录页输入框数量异常")

    inputs.nth(0).fill("")
    inputs.nth(1).fill("")
    inputs.nth(2).fill("")

    inputs.nth(0).fill(USERNAME)
    inputs.nth(1).fill(PASSWORD)
    inputs.nth(2).fill(code)


def login(page, max_retries: int = 8):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3500)

    for attempt in range(1, max_retries + 1):
        try:
            best_code = solve_captcha(page)

            if len(best_code) != 4:
                print(f"第 {attempt} 次验证码长度异常: {best_code}")
                page.goto(LOGIN_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)
                continue

            candidates = generate_code_candidates(best_code, limit=12)
            print("captcha candidates:", candidates)

            for cand in candidates:
                try:
                    fill_login_form(page, cand)
                    page.click("text=Login")
                    page.wait_for_timeout(3500)

                    body_text = page.locator("body").inner_text(timeout=5000)

                    if ("统一认证中心" not in body_text) and ("Login" not in body_text):
                        print(f"登录成功，attempt={attempt}, code={cand}")
                        return

                    print(f"候选验证码失败: {cand}")
                    page.goto(LOGIN_URL, wait_until="domcontentloaded")
                    page.wait_for_timeout(2500)

                except Exception as inner_e:
                    print(f"候选验证码异常 {cand}: {inner_e}")
                    page.goto(LOGIN_URL, wait_until="domcontentloaded")
                    page.wait_for_timeout(2500)

            print(f"第 {attempt} 次所有候选均失败，准备刷新验证码重试")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)

        except Exception as e:
            print(f"第 {attempt} 次登录异常: {e}")
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)

    raise RuntimeError("多次尝试后仍无法登录")


def open_mission_page(page):
    page.goto(MISSION_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(8000)


def expand_all(page):
    selectors = [
        "svg",
        "i",
        "button",
        "[class*='arrow']",
        "[class*='expand']",
        "[class*='toggle']",
        "[class*='icon']",
    ]

    for selector in selectors:
        try:
            items = page.locator(selector)
            for i in range(items.count()):
                try:
                    items.nth(i).click(timeout=800)
                    page.wait_for_timeout(200)
                except:
                    pass
        except:
            pass

    page.wait_for_timeout(3000)


def create_test_calendar(page):
    text = page.locator("body").inner_text()

    c = Calendar()
    e = Event()

    start = datetime.now() + timedelta(days=1)
    start = start.replace(hour=20, minute=0, second=0, microsecond=0)

    e.name = "Task Debug"
    e.begin = start
    e.end = start + timedelta(hours=1)
    e.description = text[:3000]

    c.events.add(e)

    with open("crew_schedule.ics", "w", encoding="utf-8") as f:
        f.writelines(c)


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})

        login(page, max_retries=8)
        open_mission_page(page)
        expand_all(page)
        create_test_calendar(page)

        browser.close()


if __name__ == "__main__":
    run()
