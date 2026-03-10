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

ARTIFACT_DIR = "debug_output"
os.makedirs(ARTIFACT_DIR, exist_ok=True)


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


def save_text(filename: str, text: str):
    with open(os.path.join(ARTIFACT_DIR, filename), "w", encoding="utf-8") as f:
        f.write(text)


def extract_captcha_bytes(page) -> bytes:
    imgs = page.locator("img")
    count = imgs.count()

    for i in range(count):
        try:
            src = imgs.nth(i).get_attribute("src", timeout=1000)
            if src and src.startswith("data:image"):
                return base64.b64decode(src.split(",", 1)[1])
        except Exception:
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
            variant.save(os.path.join(ARTIFACT_DIR, f"captcha_variant_{idx}.png"))
        except Exception:
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
    print("captcha candidates:", generate_code_candidates(best, limit=12))
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

            for cand in candidates:
                try:
                    fill_login_form(page, cand)
                    page.click("text=Login")
                    page.wait_for_timeout(3500)

                    body_text = page.locator("body").inner_text(timeout=5000)

                    if ("统一认证中心" not in body_text) and ("Login" not in body_text):
                        print(f"登录成功，attempt={attempt}, code={cand}")
                        page.screenshot(path=os.path.join(ARTIFACT_DIR, "after_login.png"), full_page=True)
                        save_text("after_login.txt", body_text)
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
    page.screenshot(path=os.path.join(ARTIFACT_DIR, "mission_page_before_expand.png"), full_page=True)


def expand_task_rows_only(page):
    """
    只在右侧任务区域里找展开箭头，避免点到顶部菜单“意见反馈”等图标
    """
    # 任务区在页面右半部分
    right_half = page.locator("body").locator("xpath=.//*").filter(
        has_text=re.compile(r"\d{2}月\d{2}日")
    )

    # 先尝试直接点所有含日期的行末按钮/图标
    count = right_half.count()
    print("task-like row count:", count)

    for i in range(count):
        try:
            row = right_half.nth(i)
            # 行里常见可点元素
            for selector in ["button", "svg", "i", "[class*='arrow']", "[class*='icon']"]:
                try:
                    items = row.locator(selector)
                    for j in range(items.count()):
                        try:
                            items.nth(j).click(timeout=800)
                            page.wait_for_timeout(200)
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

    page.wait_for_timeout(3000)
    page.screenshot(path=os.path.join(ARTIFACT_DIR, "mission_page_after_expand.png"), full_page=True)


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\r", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_task_area_text(page) -> str:
    text = page.locator("body").inner_text()
    text = normalize_text(text)
    save_text("mission_body_text.txt", text)

    markers = ["01月", "02月", "03月", "04月", "05月", "06月", "07月", "08月", "09月", "10月", "11月", "12月"]
    start_idx = -1
    for marker in markers:
        idx = text.find(marker)
        if idx != -1:
            start_idx = idx
            break

    if start_idx == -1:
        return text

    return text[start_idx:]


def split_tasks(task_area_text: str):
    pattern = r'(?=(\d{2}月\d{2}日\s*周.\s*(?:航班|置位|训练|摆渡|备份|待命|考勤|任务)))'
    parts = re.split(pattern, task_area_text)

    blocks = []
    current = ""

    for part in parts:
        if not part:
            continue
        if re.match(r'\d{2}月\d{2}日\s*周.', part):
            if current.strip():
                blocks.append(current.strip())
            current = part
        else:
            current += part

    if current.strip():
        blocks.append(current.strip())

    cleaned = []
    for b in blocks:
        b = normalize_text(b)
        if len(b) > 8:
            cleaned.append(b)

    save_text("task_blocks.txt", "\n\n==========\n\n".join(cleaned))
    return cleaned


def detect_task_type(block: str) -> str:
    for t in ["航班", "置位", "训练", "摆渡", "备份", "待命", "考勤"]:
        if t in block:
            return t
    return "任务"


def detect_icon(task_type: str) -> str:
    return {
        "航班": "✈️",
        "置位": "📍",
        "训练": "🎓",
        "摆渡": "🚐",
        "备份": "🗂",
        "待命": "🕒",
        "考勤": "📋",
        "任务": "🗂",
    }.get(task_type, "🗂")


def extract_flight_no(block: str) -> str:
    m = re.search(r'\b9C\d{3,4}[A-Z]?\b', block)
    return m.group(0) if m else ""


def extract_airports(block: str):
    codes = re.findall(r'\b[A-Z]{4}\b', block)
    if len(codes) >= 2:
        return codes[0], codes[1]
    return "", ""


def extract_reg_and_model(block: str):
    reg = ""
    model = ""

    m_reg = re.search(r'\bB\d{3,4}[A-Z]?\b', block)
    if m_reg:
        reg = m_reg.group(0)

    m_model = re.search(r'\b(A319|A320|A321|B300X|B321F|B737|B738|B739)\b', block)
    if m_model:
        model = m_model.group(0)

    return reg, model


def extract_times(block: str):
    times = re.findall(r'\b\d{2}:\d{2}\b', block)
    if len(times) >= 2:
        return times[0], times[-1], times
    return "", "", times


def extract_checkin_time(block: str, all_times):
    if len(all_times) >= 3:
        return all_times[-2]
    if len(all_times) == 2:
        return all_times[0]
    return ""


def extract_date(block: str):
    m = re.search(r'(\d{2})月(\d{2})日', block)
    if not m:
        return None
    month = int(m.group(1))
    day = int(m.group(2))
    year = datetime.now().year
    return year, month, day


def make_datetime(year, month, day, hhmm):
    hh, mm = map(int, hhmm.split(":"))
    return datetime(year, month, day, hh, mm)


def build_title(task_type, flight_no, dep, arr, model, reg):
    icon = detect_icon(task_type)

    if flight_no and dep and arr:
        title = f"{icon} {flight_no} {dep}→{arr}"
    elif dep and arr:
        title = f"{icon} {task_type} {dep}→{arr}"
    elif flight_no:
        title = f"{icon} {flight_no} {task_type}"
    else:
        title = f"{icon} {task_type}"

    extra = " ".join([x for x in [model, reg] if x])
    if extra:
        title += f"\n{extra}"

    return title


def build_description(block, task_type, flight_no, dep, arr, model, reg, start_time, end_time, checkin_time):
    lines = []
    lines.append(f"任务类型：{task_type}")
    if flight_no:
        lines.append(f"航班号：{flight_no}")
    if dep or arr:
        lines.append(f"航线：{dep} → {arr}")
    if model:
        lines.append(f"机型：{model}")
    if reg:
        lines.append(f"注册号：{reg}")
    if checkin_time:
        lines.append(f"签到时间：{checkin_time}")
    if start_time and end_time:
        lines.append(f"任务时间：{start_time} - {end_time}")
    lines.append("")
    lines.append("原始内容：")
    lines.append(block)
    return "\n".join(lines)


def create_calendar_from_blocks(blocks):
    c = Calendar()
    now = datetime.now()

    if not blocks:
        e = Event()
        start = now + timedelta(days=1)
        start = start.replace(hour=20, minute=0, second=0, microsecond=0)
        e.name = "🗂 未解析到任务"
        e.begin = start
        e.end = start + timedelta(hours=1)
        e.description = "脚本已运行，但未解析到任务块。请查看 debug_output。"
        c.events.add(e)
    else:
        for idx, block in enumerate(blocks):
            task_type = detect_task_type(block)
            flight_no = extract_flight_no(block)
            dep, arr = extract_airports(block)
            reg, model = extract_reg_and_model(block)
            start_time, end_time, all_times = extract_times(block)
            checkin_time = extract_checkin_time(block, all_times)
            date_info = extract_date(block)

            if date_info and start_time and end_time:
                year, month, day = date_info
                start_dt = make_datetime(year, month, day, start_time)
                end_dt = make_datetime(year, month, day, end_time)
                if end_dt <= start_dt:
                    end_dt += timedelta(days=1)
            else:
                start_dt = now + timedelta(days=1, hours=idx)
                end_dt = start_dt + timedelta(hours=1)

            e = Event()
            e.name = build_title(task_type, flight_no, dep, arr, model, reg)
            e.begin = start_dt
            e.end = end_dt
            e.description = build_description(
                block, task_type, flight_no, dep, arr, model, reg,
                start_time, end_time, checkin_time
            )
            c.events.add(e)

    with open("crew_schedule.ics", "w", encoding="utf-8") as f:
        f.writelines(c)


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})

        login(page, max_retries=8)
        open_mission_page(page)
        expand_task_rows_only(page)

        task_area_text = extract_task_area_text(page)
        blocks = split_tasks(task_area_text)
        create_calendar_from_blocks(blocks)

        browser.close()


if __name__ == "__main__":
    run()
