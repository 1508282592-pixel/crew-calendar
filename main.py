import os
import re
import io
import base64
from itertools import product
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

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

SH_TZ = ZoneInfo("Asia/Shanghai")


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
        "0": ["0", "O"], "O": ["O", "0"],
        "1": ["1", "I", "L"], "I": ["I", "1", "L"], "L": ["L", "1", "I"],
        "5": ["5", "S"], "S": ["S", "5"],
        "8": ["8", "B"], "B": ["B", "8"],
        "2": ["2", "Z"], "Z": ["Z", "2"],
        "6": ["6", "G"], "G": ["G", "6"],
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
            if cleaned:
                candidates.append(cleaned)

    if not candidates:
        return ""

    candidates = sorted(candidates, key=score_candidate, reverse=True)
    return candidates[0][:4]


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

    for _attempt in range(1, max_retries + 1):
        best_code = solve_captcha(page)

        if len(best_code) != 4:
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
                    return

                page.goto(LOGIN_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)

            except Exception:
                page.goto(LOGIN_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)

    raise RuntimeError("多次尝试后仍无法登录")


def open_mission_page(page):
    page.goto(MISSION_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(8000)


def expand_task_rows_only(page):
    page.wait_for_timeout(2000)

    body = page.locator("body")
    text = body.inner_text()

    lines = []
    for line in text.splitlines():
        line = line.strip()
        if re.search(r"\d{2}月\d{2}日\s*周.", line):
            lines.append(line)

    for line in lines:
        try:
            row = page.locator(f"text={line}").first
            box = row.bounding_box()
            if not box:
                continue

            x = box["x"] + box["width"] - 28
            y = box["y"] + box["height"] / 2
            page.mouse.click(x, y)
            page.wait_for_timeout(800)

        except Exception:
            pass

    page.wait_for_timeout(3000)


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

    markers = ["01月", "02月", "03月", "04月", "05月", "06月",
               "07月", "08月", "09月", "10月", "11月", "12月"]
    start_idx = -1
    for marker in markers:
        idx = text.find(marker)
        if idx != -1:
            start_idx = idx
            break

    if start_idx == -1:
        return text

    return text[start_idx:]


def split_day_blocks(task_area_text: str):
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

    out = [normalize_text(b) for b in blocks if len(normalize_text(b)) > 8]
    save_text("task_blocks.txt", "\n\n==========\n\n".join(out))
    return out


def detect_task_type(day_block: str) -> str:
    header = day_block.splitlines()[0] if day_block.splitlines() else day_block
    for t in ["置位", "航班", "训练", "摆渡", "备份", "待命", "考勤"]:
        if t in header:
            return t
    return "任务"


def task_bucket(task_type: str) -> str:
    return {
        "航班": "flight",
        "置位": "positioning",
        "训练": "training",
        "摆渡": "ferry",
    }.get(task_type, "other")


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


def extract_date(day_block: str):
    m = re.search(r'(\d{2})月(\d{2})日', day_block)
    if not m:
        return None
    return datetime.now(SH_TZ).year, int(m.group(1)), int(m.group(2))


def make_datetime(year, month, day, hhmm):
    hh, mm = map(int, hhmm.split(":"))
    return datetime(year, month, day, hh, mm, tzinfo=SH_TZ)


def split_segments(day_block: str):
    positions = list(re.finditer(r'\b9C\d{3,4}[A-Z]?\b', day_block))
    if not positions:
        return [day_block]

    segments = []
    for i, m in enumerate(positions):
        start = m.start()
        end = positions[i + 1].start() if i + 1 < len(positions) else len(day_block)
        seg = day_block[start:end].strip()
        if len(seg) > 5:
            segments.append(seg)

    return segments or [day_block]


def extract_flight_no(segment: str) -> str:
    m = re.search(r'\b9C\d{3,4}[A-Z]?\b', segment)
    return m.group(0) if m else ""


def extract_airports(segment: str):
    codes = re.findall(r'\b[A-Z]{4}\b', segment)
    uniq = []
    for c in codes:
        if c not in uniq:
            uniq.append(c)
    if len(uniq) >= 2:
        return uniq[0], uniq[1]
    return "", ""


def extract_reg_and_model(segment: str):
    reg = ""
    model = ""

    m_combo = re.search(r'(B\d{3,4}[A-Z]?)(A3\d{2})', segment)
    if m_combo:
        return m_combo.group(1), m_combo.group(2)

    m_reg = re.search(r'\bB\d{3,4}[A-Z]?\b', segment)
    if m_reg:
        reg = m_reg.group(0)

    m_model = re.search(r'\b(A319|A320|A321|B300X|B321F|B737|B738|B739)\b', segment)
    if m_model:
        model = m_model.group(0)

    return reg, model


def extract_start_end_time(segment: str):
    m = re.search(r'([^\s]+)[^\n]*?(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})', segment)
    if m:
        return m.group(2), m.group(3)

    times = re.findall(r'\b\d{2}:\d{2}\b', segment)
    if len(times) >= 2:
        return times[-2], times[-1]

    return "", ""


def extract_checkin(segment: str):
    m = re.search(r'(\d{2}:\d{2})\s*([^\s]+)\s*航班动态', segment)
    if m:
        return m.group(1), m.group(2)
    return "", ""


def extract_people_type(segment: str):
    for t in ["随机人员", "乘务长", "副驾驶", "机长"]:
        if t in segment:
            return t
    return ""


def extract_people_lines(segment: str):
    lines = [x.strip() for x in segment.splitlines() if x.strip()]
    out = []

    capture = False
    for line in lines:
        if line in ["随机人员", "乘务长", "副驾驶", "机长"]:
            capture = True
            continue

        if not capture:
            continue

        if "航班动态" in line:
            continue
        if re.search(r'\b9C\d', line):
            continue
        if re.search(r'\b[A-Z]{4}\b', line):
            continue
        if re.search(r'\d{2}:\d{2}', line):
            continue
        if re.search(r'^B\d{3,4}', line):
            continue

        out.append(line)

    return out


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


def build_description(day_header, task_type, flight_no, dep, arr, model, reg,
                      start_time, end_time, checkin_time, checkin_place,
                      people_type, people_lines, segment):
    lines = []
    lines.append(f"日期：{day_header}")
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
    if checkin_place:
        lines.append(f"签到地点：{checkin_place}")
    if start_time and end_time:
        lines.append(f"任务时间：{start_time} - {end_time}")

    if people_type:
        lines.append(f"人员类型：{people_type}")
    if people_lines:
        lines.append("人员名单：")
        for p in people_lines:
            lines.append(p)

    lines.append("")
    lines.append("原始内容：")
    lines.append(segment)
    return "\n".join(lines)


def create_multi_calendars(day_blocks):
    calendars = {
        "flight": Calendar(),
        "positioning": Calendar(),
        "training": Calendar(),
        "ferry": Calendar(),
        "other": Calendar(),
    }

    seen = set()

    for day_block in day_blocks:
        task_type = detect_task_type(day_block)
        date_info = extract_date(day_block)

        header_line = day_block.splitlines()[0] if day_block.splitlines() else ""
        day_header = header_line

        segments = split_segments(day_block)

        for seg in segments:
            flight_no = extract_flight_no(seg)
            dep, arr = extract_airports(seg)
            reg, model = extract_reg_and_model(seg)
            start_time, end_time = extract_start_end_time(seg)
            checkin_time, checkin_place = extract_checkin(seg)
            people_type = extract_people_type(seg)
            people_lines = extract_people_lines(seg)

            if not (date_info and start_time and end_time):
                continue

            year, month, day = date_info
            start_dt = make_datetime(year, month, day, start_time)
            end_dt = make_datetime(year, month, day, end_time)

            if end_dt <= start_dt:
                end_dt += timedelta(days=1)

            if not flight_no and not (dep and arr):
                continue

            dedup_key = (
                task_type,
                flight_no,
                dep,
                arr,
                reg,
                model,
                start_dt.isoformat(),
                end_dt.isoformat(),
            )
            if dedup_key in seen:
                continue
            seen.add(dedup_key)

            e = Event()
            e.name = build_title(task_type, flight_no, dep, arr, model, reg)
            e.begin = start_dt
            e.end = end_dt
            e.description = build_description(
                day_header, task_type, flight_no, dep, arr, model, reg,
                start_time, end_time, checkin_time, checkin_place,
                people_type, people_lines, seg
            )

            bucket = task_bucket(task_type)
            calendars[bucket].events.add(e)

    mapping = {
        "flight": "flight.ics",
        "positioning": "positioning.ics",
        "training": "training.ics",
        "ferry": "ferry.ics",
        "other": "other.ics",
    }

    for key, filename in mapping.items():
        with open(filename, "w", encoding="utf-8") as f:
            f.writelines(calendars[key])


def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})

        login(page, max_retries=8)
        open_mission_page(page)
        expand_task_rows_only(page)

        task_area_text = extract_task_area_text(page)
        day_blocks = split_day_blocks(task_area_text)
        create_multi_calendars(day_blocks)

        browser.close()


if __name__ == "__main__":
    run()
