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

    for variant in variants:
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


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\r", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_date(day_block: str):
    m = re.search(r'(\d{2})月(\d{2})日', day_block)
    if not m:
        return None
    return datetime.now(SH_TZ).year, int(m.group(1)), int(m.group(2))


def make_datetime(year, month, day, hhmm):
    hh, mm = map(int, hhmm.split(":"))
    return datetime(year, month, day, hh, mm, tzinfo=SH_TZ)


def detect_task_type(text: str) -> str:
    for t in ["置位", "航班", "训练", "摆渡", "备份", "待命", "考勤"]:
        if t in text:
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
    m_combo = re.search(r'(B\d{3,4}[A-Z]?)(A3\d{2})', segment)
    if m_combo:
        return m_combo.group(1), m_combo.group(2)

    reg = ""
    model = ""

    m_reg = re.search(r'\bB\d{3,4}[A-Z]?\b', segment)
    if m_reg:
        reg = m_reg.group(0)

    m_model = re.search(r'\b(A319|A320|A321|B300X|B321F|B737|B738|B739)\b', segment)
    if m_model:
        model = m_model.group(0)

    return reg, model


def extract_start_end_time(segment: str):
    m = re.search(r'(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})', segment)
    if m:
        return m.group(1), m.group(2)

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
        if "查看更多" in line:
            continue
        if re.match(r'\d{4}-\d{2}-\d{2}', line):
            continue

        out.append(line)

    return out


def build_title(task_type, flight_no, dep, arr, model, reg):
    icon = detect_icon(task_type)

    if flight_no and dep and arr:
        title = f"{icon} {flight_no} {dep}→{arr}"
    elif flight_no:
        title = f"{icon} {flight_no}"
    elif dep and arr:
        title = f"{icon} {dep}→{arr}"
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


def event_quality(flight_no, dep, arr, reg, model, checkin_time, checkin_place, people_lines):
    score = 0
    if flight_no:
        score += 10
    if dep and arr:
        score += 50
    if reg:
        score += 10
    if model:
        score += 10
    if checkin_time:
        score += 10
    if checkin_place:
        score += 10
    if people_lines:
        score += 10
    return score


def get_task_header_lines(page):
    body = page.locator("body")
    text = body.inner_text()
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if re.search(r"\d{2}月\d{2}日\s*周.", line):
            lines.append(line)
    # 去重保持顺序
    seen = set()
    out = []
    for x in lines:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def collect_expanded_tasks_one_by_one(page):
    """
    逐条展开 -> 读取 -> 收起
    避免多个任务同时展开导致内容混在一起
    """
    task_entries = []
    header_lines = get_task_header_lines(page)

    for idx, header in enumerate(header_lines):
        try:
            row = page.locator(f"text={header}").first
            box = row.bounding_box()
            if not box:
                continue

            x = box["x"] + box["width"] - 28
            y = box["y"] + box["height"] / 2

            # 展开
            page.mouse.click(x, y)
            page.wait_for_timeout(1200)

            full_text = normalize_text(page.locator("body").inner_text())

            # 从当前header开始，截到下一个header前
            start = full_text.find(header)
            if start == -1:
                # 收起后继续
                page.mouse.click(x, y)
                page.wait_for_timeout(800)
                continue

            if idx + 1 < len(header_lines):
                next_header = header_lines[idx + 1]
                end = full_text.find(next_header, start + len(header))
                block = full_text[start:end].strip() if end != -1 else full_text[start:].strip()
            else:
                block = full_text[start:].strip()

            task_entries.append(block)

            # 收起
            page.mouse.click(x, y)
            page.wait_for_timeout(800)

        except Exception:
            pass

    save_text("task_entries.txt", "\n\n==========\n\n".join(task_entries))
    return task_entries


def create_multi_calendars(task_entries):
    calendars = {
        "flight": Calendar(),
        "positioning": Calendar(),
        "training": Calendar(),
        "ferry": Calendar(),
        "other": Calendar(),
    }

    best_events = {}

    for entry in task_entries:
        task_type = detect_task_type(entry)
        date_info = extract_date(entry)
        header_line = entry.splitlines()[0] if entry.splitlines() else ""
        day_header = header_line

        flight_no = extract_flight_no(entry)
        dep, arr = extract_airports(entry)
        reg, model = extract_reg_and_model(entry)
        start_time, end_time = extract_start_end_time(entry)
        checkin_time, checkin_place = extract_checkin(entry)
        people_type = extract_people_type(entry)
        people_lines = extract_people_lines(entry)

        if not date_info or not flight_no or not start_time or not end_time:
            continue

        year, month, day = date_info
        start_dt = make_datetime(year, month, day, start_time)
        end_dt = make_datetime(year, month, day, end_time)

        if end_dt <= start_dt:
            end_dt += timedelta(days=1)

        group_key = (
            task_type,
            flight_no,
            start_dt.isoformat(),
            end_dt.isoformat(),
        )

        quality = event_quality(
            flight_no, dep, arr, reg, model, checkin_time, checkin_place, people_lines
        )

        candidate = {
            "task_type": task_type,
            "flight_no": flight_no,
            "dep": dep,
            "arr": arr,
            "reg": reg,
            "model": model,
            "start_dt": start_dt,
            "end_dt": end_dt,
            "start_time": start_time,
            "end_time": end_time,
            "checkin_time": checkin_time,
            "checkin_place": checkin_place,
            "people_type": people_type,
            "people_lines": people_lines,
            "day_header": day_header,
            "entry": entry,
            "quality": quality,
        }

        if group_key not in best_events or quality > best_events[group_key]["quality"]:
            best_events[group_key] = candidate

    for item in best_events.values():
        e = Event()
        e.name = build_title(
            item["task_type"], item["flight_no"], item["dep"], item["arr"],
            item["model"], item["reg"]
        )
        e.begin = item["start_dt"]
        e.end = item["end_dt"]
        e.description = build_description(
            item["day_header"], item["task_type"], item["flight_no"],
            item["dep"], item["arr"], item["model"], item["reg"],
            item["start_time"], item["end_time"], item["checkin_time"],
            item["checkin_place"], item["people_type"], item["people_lines"],
            item["entry"]
        )

        bucket = task_bucket(item["task_type"])
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

        task_entries = collect_expanded_tasks_one_by_one(page)
        create_multi_calendars(task_entries)

        browser.close()


if __name__ == "__main__":
    run()
