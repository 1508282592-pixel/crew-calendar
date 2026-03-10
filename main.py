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

LOGIN_URL = "https://cp.9cair.com"
MISSION_URL = "https://cp.9cair.com/html/task/mission.html"

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]

ARTIFACT_DIR = "debug_output"
os.makedirs(ARTIFACT_DIR, exist_ok=True)

SH_TZ = ZoneInfo("Asia/Shanghai")

AIRPORT_CN_TO_ICAO = {
    "上海虹桥": "ZSSS",
    "上海浦东": "ZSPD",
    "西安咸阳": "ZLXY",
    "重庆江北": "ZUCK",
    "大连周水子": "ZYTL",
    "深圳宝安": "ZGSZ",
    "济南遥墙": "ZSJN",
    "哈尔滨太平": "ZYHB",
    "淮安涟水": "ZSSH",
    "呼和浩特白塔": "ZBHH",
    "长春龙嘉": "ZYCC",
    "兰州中川": "ZLLL",
}

AIRPORT_NAMES = sorted(AIRPORT_CN_TO_ICAO.keys(), key=len, reverse=True)


def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\r", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def save_text(filename: str, text: str):
    with open(os.path.join(ARTIFACT_DIR, filename), "w", encoding="utf-8") as f:
        f.write(text)


def escape_ics_text(text: str) -> str:
    text = text.replace("\\", "\\\\")
    text = text.replace(";", r"\;")
    text = text.replace(",", r"\,")
    text = text.replace("\n", r"\n")
    return text


def format_dt_local(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")


def fr24_flight_code(flight_number: str) -> str:
    return re.sub(r"[A-Za-z]$", "", flight_number.strip())


def make_datetime(year: int, month: int, day: int, hhmm: str) -> datetime:
    hh, mm = map(int, hhmm.split(":"))
    return datetime(year, month, day, hh, mm, tzinfo=SH_TZ)


# =========================
# 验证码
# =========================

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


# =========================
# 登录
# =========================

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
    for _attempt in range(1, max_retries + 1):
        try:
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(5000)
        except Exception:
            if _attempt == max_retries:
                raise
            page.wait_for_timeout(5000)
            continue

        best_code = solve_captcha(page)
        if len(best_code) != 4:
            page.wait_for_timeout(2000)
            continue

        candidates = generate_code_candidates(best_code, limit=12)

        for cand in candidates:
            try:
                fill_login_form(page, cand)
                page.click("text=Login")
                page.wait_for_timeout(4000)

                body_text = page.locator("body").inner_text(timeout=8000)
                if ("统一认证中心" not in body_text) and ("Login" not in body_text):
                    return

                page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90000)
                page.wait_for_timeout(3000)

            except Exception:
                try:
                    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90000)
                    page.wait_for_timeout(3000)
                except Exception:
                    pass

    raise RuntimeError("多次尝试后仍无法登录")


# =========================
# 页面采集
# =========================

def open_mission_page(page):
    for i in range(3):
        try:
            page.goto(MISSION_URL, wait_until="domcontentloaded", timeout=90000)
            page.wait_for_timeout(8000)
            return
        except Exception:
            if i == 2:
                raise
            page.wait_for_timeout(5000)


def get_day_headers(page):
    text = page.locator("body").inner_text()
    headers = []
    for line in text.splitlines():
        line = line.strip()
        if re.search(r"\d{2}月\d{2}日\s*周.", line):
            headers.append(line)

    seen = set()
    out = []
    for h in headers:
        if h not in seen:
            seen.add(h)
            out.append(h)
    return out


def click_day_toggle(page, header: str):
    row = page.locator(f"text={header}").first
    box = row.bounding_box()
    if not box:
        return False

    x = box["x"] + box["width"] - 28
    y = box["y"] + box["height"] / 2
    page.mouse.click(x, y)
    return True


def expand_day(page, header: str):
    ok = click_day_toggle(page, header)
    if ok:
        page.wait_for_timeout(1500)
    return ok


def collapse_day(page, header: str):
    try:
        ok = click_day_toggle(page, header)
        if ok:
            page.wait_for_timeout(800)
    except Exception:
        pass


# =========================
# 分段：按详细卡头切
# =========================

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


def extract_date(text: str):
    m = re.search(r'(\d{2})月(\d{2})日', text)
    if not m:
        return None
    return datetime.now(SH_TZ).year, int(m.group(1)), int(m.group(2))


def is_card_header_line(line: str) -> bool:
    line = normalize_text(line)
    patterns = [
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+A3\d{2}$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+\|\s+A3\d{2}$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+A320$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+A321$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+A319$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+\|\s+A320$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+\|\s+A321$',
        r'^9C\d{3,4}[A-Z]?\s+B\d{3,4}[A-Z]{0,2}\s+\|\s+A319$',
    ]
    return any(re.fullmatch(p, line) for p in patterns)


def split_day_block_by_card_headers(day_block: str):
    lines = [normalize_text(x) for x in day_block.splitlines() if normalize_text(x)]
    if not lines:
        return []

    starts = []
    for i, line in enumerate(lines):
        if is_card_header_line(line):
            starts.append(i)

    cards = []
    for idx, start_i in enumerate(starts):
        end_i = starts[idx + 1] if idx + 1 < len(starts) else len(lines)
        chunk_lines = lines[start_i:end_i]
        chunk = "\n".join(chunk_lines).strip()

        if "航班动态" not in chunk:
            continue
        if not re.search(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', chunk):
            continue
        if not re.search(r'\b9C\d{3,4}[A-Z]?\b', chunk):
            continue

        cards.append(chunk)

    uniq = []
    seen = set()
    for c in cards:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


def collect_day_blocks(page):
    day_headers = get_day_headers(page)
    save_text("day_headers.txt", "\n".join(day_headers))

    result = []

    for idx, header in enumerate(day_headers):
        if not expand_day(page, header):
            continue

        try:
            body_text = normalize_text(page.locator("body").inner_text())
            start = body_text.find(header)
            if start == -1:
                continue

            if idx + 1 < len(day_headers):
                next_header = day_headers[idx + 1]
                end = body_text.find(next_header, start + len(header))
                day_block = body_text[start:end].strip() if end != -1 else body_text[start:].strip()
            else:
                day_block = body_text[start:].strip()

            result.append({
                "day_header": header,
                "task_type": detect_task_type(header),
                "day_block": day_block,
                "cards": split_day_block_by_card_headers(day_block),
            })

        finally:
            collapse_day(page, header)

    save_text(
        "day_blocks_debug.txt",
        "\n\n========== DAY ==========\n\n".join(
            f"{x['day_header']}\n\n{x['day_block']}" for x in result
        )
    )

    save_text(
        "cards_debug.txt",
        "\n\n========== DAY ==========\n\n".join(
            f"{x['day_header']}\n\n" + "\n\n------ CARD ------\n\n".join(x["cards"])
            for x in result
        )
    )

    return result


# =========================
# 卡片解析
# =========================

def extract_flight_no(card_text: str) -> str:
    m = re.search(r'\b9C\d{3,4}[A-Z]?\b', card_text)
    return m.group(0) if m else ""


def extract_reg_and_model(card_text: str):
    patterns = [
        r'\b(B\d{3,4}[A-Z]{0,2})\s+\|\s+(A319|A320|A321)\b',
        r'\b(B\d{3,4}[A-Z]{0,2})\s+(A319|A320|A321)\b',
        r'\b(B\d{3,4}[A-Z]{0,2})(A319|A320|A321)\b',
    ]
    for p in patterns:
        m = re.search(p, card_text)
        if m:
            return m.group(1), m.group(2)
    return "", ""


def extract_checkin(card_text: str):
    m = re.search(r'(\d{2}:\d{2})\s*([^\s]+)\s*航班动态', card_text)
    if m:
        return m.group(1), m.group(2)
    return "", ""


def extract_start_end_time(card_text: str):
    ranges = re.findall(r'(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})', card_text)
    if ranges:
        return ranges[-1][0], ranges[-1][1]
    return "", ""


def parse_route_cn_from_line(line: str):
    line = re.sub(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', '', line).strip()
    line = line.replace("→", "").replace("-", "").replace("—", "")
    for dep_cn in AIRPORT_NAMES:
        if line.startswith(dep_cn):
            remain = line[len(dep_cn):].strip()
            for arr_cn in AIRPORT_NAMES:
                if remain == arr_cn:
                    return dep_cn, arr_cn
    return "", ""


def extract_airports(card_text: str):
    dep_cn = ""
    arr_cn = ""

    lines = [normalize_text(x) for x in card_text.splitlines() if normalize_text(x)]
    candidate_lines = []

    for line in lines:
        if re.search(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', line) and "航班动态" not in line:
            candidate_lines.append(line)

    if candidate_lines:
        dep_cn, arr_cn = parse_route_cn_from_line(candidate_lines[-1])

    dep = AIRPORT_CN_TO_ICAO.get(dep_cn, "")
    arr = AIRPORT_CN_TO_ICAO.get(arr_cn, "")

    if dep and arr:
        return dep, arr, dep_cn, arr_cn

    codes = re.findall(r'\b[A-Z]{4}\b', card_text)
    uniq = []
    for c in codes:
        if c not in uniq:
            uniq.append(c)
    if len(uniq) >= 2:
        return uniq[0], uniq[1], dep_cn, arr_cn)

    return "", "", dep_cn, arr_cn


def extract_people_lines(card_text: str):
    lines = [normalize_text(x) for x in card_text.splitlines() if normalize_text(x)]
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
        if re.search(r'\b9C\d{3,4}[A-Z]?\b', line):
            continue
        if re.search(r'\d{2}:\d{2}', line):
            continue
        if re.search(r'^B\d{3,4}', line):
            continue
        if re.search(r'\b(A319|A320|A321)\b', line):
            continue
        if "查看更多" in line:
            continue
        if re.match(r'\d{4}-\d{2}-\d{2}', line):
            continue

        out.append(line)

    uniq = []
    seen = set()
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


# =========================
# ICS 输出
# =========================

def build_title(task_type, flight_no, dep, arr, dep_cn, arr_cn):
    icon = detect_icon(task_type)

    if flight_no and dep and arr:
        return f"{icon} {flight_no} {dep}→{arr}"
    if flight_no and dep_cn and arr_cn:
        return f"{icon} {flight_no} {dep_cn}→{arr_cn}"
    if flight_no:
        return f"{icon} {flight_no}"
    return f"{icon} {task_type}"


def build_description(item: dict) -> str:
    fr24_number = fr24_flight_code(item["flight_no"]) if item["flight_no"] else ""

    lines = []
    lines.append(f"{item['day_header']}")
    lines.append(f"航班号：{item['flight_no']}")

    if item["dep"] or item["arr"]:
        lines.append(f"航线：{item['dep']} → {item['arr']}")
    if item["dep_cn"] or item["arr_cn"]:
        lines.append(f"航线：{item['dep_cn']} → {item['arr_cn']}")

    if item["checkin_time"]:
        lines.append(f"签到时间：{item['checkin_time']}")
    if item["checkin_place"]:
        lines.append(f"签到地点：{item['checkin_place']}")

    lines.append(f"任务时间：{item['start_time']} - {item['end_time']}")

    if item["model"]:
        lines.append(f"机型：{item['model']}")
    if item["reg"]:
        lines.append(f"注册号：{item['reg']}")

    if item["people_lines"]:
        lines.append("")
        lines.append("人员名单：")
        for p in item["people_lines"]:
            lines.append(p)

    if fr24_number:
        lines.append("")
        lines.append(f"航班追踪：https://www.flightradar24.com/data/flights/{fr24_number}")

    if item["reg"]:
        lines.append(f"机号信息：https://www.flightradar24.com/data/aircraft/{item['reg']}")

    return "\n".join(lines)


def build_vevent(item: dict) -> str:
    title = build_title(
        item["task_type"], item["flight_no"], item["dep"], item["arr"],
        item["dep_cn"], item["arr_cn"]
    )
    desc = build_description(item)

    uid = (
        f'{item["task_type"]}-'
        f'{item["flight_no"]}-'
        f'{format_dt_local(item["start_dt"])}-'
        f'{format_dt_local(item["end_dt"])}@crew-calendar'
    )

    return "\n".join([
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"SUMMARY:{escape_ics_text(title)}",
        f"DTSTART;TZID=Asia/Shanghai:{format_dt_local(item['start_dt'])}",
        f"DTEND;TZID=Asia/Shanghai:{format_dt_local(item['end_dt'])}",
        f"DESCRIPTION:{escape_ics_text(desc)}",
        "BEGIN:VALARM",
        "TRIGGER:-PT90M",
        "DESCRIPTION:签到提醒",
        "ACTION:DISPLAY",
        "END:VALARM",
        "END:VEVENT",
    ])


def write_calendar(filename: str, items: list[dict]):
    content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Crew Calendar//CN",
    ]
    for item in items:
        content.append(build_vevent(item))
    content.append("END:VCALENDAR")

    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(content))


def event_quality(item: dict) -> int:
    score = 0
    if item["flight_no"]:
        score += 10
    if item["dep"] and item["arr"]:
        score += 50
    if item["dep_cn"] and item["arr_cn"]:
        score += 20
    if item["reg"]:
        score += 10
    if item["model"]:
        score += 10
    if item["checkin_time"]:
        score += 10
    if item["checkin_place"]:
        score += 10
    if item["people_lines"]:
        score += 10
    return score


# =========================
# 生成日历
# =========================

def create_multi_calendars_from_blocks(day_blocks):
    buckets = {
        "flight": [],
        "positioning": [],
        "training": [],
        "ferry": [],
        "other": [],
    }
    best_events = {}

    for day in day_blocks:
        date_info = extract_date(day["day_header"])
        if not date_info:
            continue

        year, month, day_num = date_info
        task_type = day["task_type"]

        for card_text in day["cards"]:
            flight_no = extract_flight_no(card_text)
            reg, model = extract_reg_and_model(card_text)
            start_time, end_time = extract_start_end_time(card_text)
            checkin_time, checkin_place = extract_checkin(card_text)
            dep, arr, dep_cn, arr_cn = extract_airports(card_text)
            people_lines = extract_people_lines(card_text)

            if not flight_no or not start_time or not end_time:
                continue

            start_dt = make_datetime(year, month, day_num, start_time)
            end_dt = make_datetime(year, month, day_num, end_time)
            if end_dt <= start_dt:
                end_dt += timedelta(days=1)

            item = {
                "day_header": day["day_header"],
                "task_type": task_type,
                "flight_no": flight_no,
                "dep": dep,
                "arr": arr,
                "dep_cn": dep_cn,
                "arr_cn": arr_cn,
                "start_time": start_time,
                "end_time": end_time,
                "checkin_time": checkin_time,
                "checkin_place": checkin_place,
                "model": model,
                "reg": reg,
                "people_lines": people_lines,
                "start_dt": start_dt,
                "end_dt": end_dt,
            }

            group_key = (
                task_type,
                flight_no,
                start_dt.isoformat(),
                end_dt.isoformat(),
            )

            q = event_quality(item)
            item["quality"] = q

            if group_key not in best_events or q > best_events[group_key]["quality"]:
                best_events[group_key] = item

    for item in best_events.values():
        buckets[task_bucket(item["task_type"])].append(item)

    for key in buckets:
        buckets[key].sort(key=lambda x: (x["start_dt"], x["flight_no"]))

    write_calendar("flight.ics", buckets["flight"])
    write_calendar("positioning.ics", buckets["positioning"])
    write_calendar("training.ics", buckets["training"])
    write_calendar("ferry.ics", buckets["ferry"])
    write_calendar("other.ics", buckets["other"])


# =========================
# 主流程
# =========================

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(
            viewport={"width": 1400, "height": 1000},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()
        page.set_default_timeout(90000)
        page.set_default_navigation_timeout(90000)

        login(page, max_retries=8)
        open_mission_page(page)

        day_blocks = collect_day_blocks(page)
        create_multi_calendars_from_blocks(day_blocks)

        context.close()
        browser.close()


if __name__ == "__main__":
    run()
