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
}

AIRPORT_NAMES = sorted(AIRPORT_CN_TO_ICAO.keys(), key=len, reverse=True)


# =========================
# 基础工具
# =========================

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


# =========================
# 页面采集：逐天展开、逐天收起
# =========================

def open_mission_page(page):
    page.goto(MISSION_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(8000)


def get_task_header_lines(page):
    body = page.locator("body")
    text = body.inner_text()
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if re.search(r"\d{2}月\d{2}日\s*周.", line):
            lines.append(line)

    seen = set()
    out = []
    for x in lines:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def collect_day_entries_one_by_one(page):
    """
    一天一条展开，一天一条收起。
    只取当前天自己的块。
    """
    day_entries = []
    header_lines = get_task_header_lines(page)

    for idx, header in enumerate(header_lines):
        try:
            row = page.locator(f"text={header}").first
            box = row.bounding_box()
            if not box:
                continue

            x = box["x"] + box["width"] - 28
            y = box["y"] + box["height"] / 2

            page.mouse.click(x, y)
            page.wait_for_timeout(1200)

            full_text = normalize_text(page.locator("body").inner_text())
            start = full_text.find(header)
            if start == -1:
                page.mouse.click(x, y)
                page.wait_for_timeout(800)
                continue

            if idx + 1 < len(header_lines):
                next_header = header_lines[idx + 1]
                end = full_text.find(next_header, start + len(header))
                block = full_text[start:end].strip() if end != -1 else full_text[start:].strip()
            else:
                block = full_text[start:].strip()

            day_entries.append(block)

            page.mouse.click(x, y)
            page.wait_for_timeout(800)

        except Exception:
            pass

    save_text("day_entries.txt", "\n\n==========\n\n".join(day_entries))
    return day_entries


# =========================
# 解析
# =========================

def detect_task_type(text: str) -> str:
    for t in ["置位", "航班", "训练", "摆渡", "备份", "待命", "考勤"]:
        if t in text:
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


def task_bucket(task_type: str) -> str:
    return {
        "航班": "flight",
        "置位": "positioning",
        "训练": "training",
        "摆渡": "ferry",
    }.get(task_type, "other")


def extract_date(text: str):
    m = re.search(r'(\d{2})月(\d{2})日', text)
    if not m:
        return None
    return datetime.now(SH_TZ).year, int(m.group(1)), int(m.group(2))


def split_day_entry_into_detailed_segments(day_entry: str):
    """
    只拆出真正详细航段。
    详细航段必须至少满足：
    - 以独立 9Cxxxx 行开始
    - 段内有 航班动态
    - 段内有 B注册号+机型
    - 段内有一条 xx:xx-xx:xx
    这样顶部摘要里的裸航班号就不会被当成事件。
    """
    lines = [x.strip() for x in day_entry.splitlines() if x.strip()]
    if not lines:
        return []

    header = lines[0]
    segments = []
    current = None

    for line in lines[1:]:
        if re.fullmatch(r'9C\d{3,4}[A-Z]?', line):
            if current:
                segments.append(current)
            current = [header, line]
        else:
            if current is not None:
                current.append(line)

    if current:
        segments.append(current)

    out = []
    for seg_lines in segments:
        seg_text = "\n".join(seg_lines)
        has_flight_dynamic = "航班动态" in seg_text
        has_reg_model = re.search(r'B\d{3,4}[A-Z]?A3\d{2}', seg_text) is not None
        has_range = re.search(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', seg_text) is not None

        if has_flight_dynamic and has_reg_model and has_range:
            out.append(seg_text)

    return out


def extract_flight_no(segment: str) -> str:
    for line in segment.splitlines():
        line = line.strip()
        if re.fullmatch(r'9C\d{3,4}[A-Z]?', line):
            return line
    m = re.search(r'\b9C\d{3,4}[A-Z]?\b', segment)
    return m.group(0) if m else ""


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


def extract_checkin(segment: str):
    m = re.search(r'(\d{2}:\d{2})\s*([^\s]+)\s*航班动态', segment)
    if m:
        return m.group(1), m.group(2)
    return "", ""


def extract_start_end_time(segment: str):
    """
    取当前航段自己的时间区间。
    优先取最后一个时间区间，因为详细段最后那组通常就是本航段。
    """
    ranges = re.findall(r'(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})', segment)
    if ranges:
        return ranges[-1][0], ranges[-1][1]
    return "", ""


def parse_route_cn_from_line(line: str):
    line = re.sub(r'\d{2}:\d{2}\s*-\s*\d{2}:\d{2}', '', line).strip()
    for dep_cn in AIRPORT_NAMES:
        if line.startswith(dep_cn):
            remain = line[len(dep_cn):]
            for arr_cn in AIRPORT_NAMES:
                if remain == arr_cn:
                    return dep_cn, arr_cn
    return "", ""


def extract_airports(segment: str):
    """
    优先从中文航线行里取出当前航段自己的起终点。
    再映射成 ICAO。
    """
    dep_cn = ""
    arr_cn = ""

    lines = [x.strip() for x in segment.splitlines() if x.strip()]
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

    # 兜底：直接抓 ICAO
    codes = re.findall(r'\b[A-Z]{4}\b', segment)
    uniq = []
    for c in codes:
        if c not in uniq:
            uniq.append(c)
    if len(uniq) >= 2:
        return uniq[0], uniq[1], dep_cn, arr_cn

    return "", "", dep_cn, arr_cn


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
        if re.fullmatch(r'9C\d{3,4}[A-Z]?', line):
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

    seen = set()
    uniq = []
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


# =========================
# 事件模板
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
    lines.append(f"日期：{item['day_header']}")
    lines.append(f"任务类型：{item['task_type']}")
    lines.append(f"航班号：{item['flight_no']}")

    if item["dep"] or item["arr"]:
        lines.append(f"航线：{item['dep']} → {item['arr']}")
    if item["dep_cn"] or item["arr_cn"]:
        lines.append(f"中文航线：{item['dep_cn']} → {item['arr_cn']}")

    if item["checkin_time"]:
        lines.append(f"签到时间：{item['checkin_time']}")
    if item["checkin_place"]:
        lines.append(f"签到地点：{item['checkin_place']}")

    lines.append(f"任务时间：{item['start_time']} - {item['end_time']}")

    if item["model"]:
        lines.append(f"机型：{item['model']}")
    if item["reg"]:
        lines.append(f"注册号：{item['reg']}")

    if item["people_type"]:
        lines.append("")
        lines.append(f"人员类型：{item['people_type']}")

    if item["people_lines"]:
        lines.append("人员名单：")
        for p in item["people_lines"]:
            lines.append(p)

    if fr24_number:
        lines.append("")
        lines.append(f"航班追踪：https://www.flightradar24.com/data/flights/{fr24_number}")
    if item["reg"]:
        lines.append(f"机号信息：https://www.flightradar24.com/data/
