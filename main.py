import sys
sys.stdout.reconfigure(line_buffering=True)

import os
import re
import json
import hmac
import base64
import hashlib
import datetime
import threading

import requests as req
import gspread
from flask import Flask, request
from google.oauth2.service_account import Credentials
from openai import OpenAI

app = Flask(__name__)

# ===== 環境變數：請在 Render 設定，不要把金鑰寫進 GitHub =====
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
TW = datetime.timezone(datetime.timedelta(hours=8))

# ===== 大項目與小項目 =====
# 支出保留原本的小項目，另外新增「交卡費」。
# 刷卡的小項目與支出完全相同（包含「交卡費」）。
EXPENSE_CATEGORIES = [
    "餐飲", "交通", "生活", "醫療健康", "美容保養", "服飾",
    "訂閱服務", "娛樂休閒", "運動健身", "旅遊", "教育", "保險",
    "稅費", "捐款公益", "其他", "交卡費"
]
CARD_CATEGORIES = EXPENSE_CATEGORIES.copy()
INCOME_CATEGORIES = [
    "薪資", "獎金", "投資收入", "利息", "退款", "紅包",
    "補助", "兼職收入", "其他"
]
MAIN_TYPES = ["支出", "收入", "刷卡"]

HEADERS = ["日期", "時間", "傳送者", "大項目", "項目", "金額", "小項目", "原始訊息"]


def now_tw():
    return datetime.datetime.now(TW)


def get_sheet():
    """連接試算表；若是舊版七欄格式，會自動轉成新版八欄格式並保留舊資料。"""
    if not GOOGLE_CREDENTIALS_JSON:
        raise RuntimeError("Render 尚未設定 GOOGLE_CREDENTIALS_JSON")
    if not GOOGLE_SHEET_ID:
        raise RuntimeError("Render 尚未設定 GOOGLE_SHEET_ID")

    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)

    try:
        ws = sh.worksheet("記帳")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="記帳", rows="2000", cols="8")
        ws.update("A1", [HEADERS])
        return ws

    values = ws.get_all_values()
    if not values:
        ws.update("A1", [HEADERS])
        return ws

    if values[0] != HEADERS:
        old_headers = values[0]
        # 舊版欄位：日期、時間、傳送者、項目、金額、類別、原始訊息
        old_rows = values[1:]
        if old_headers[:7] == ["日期", "時間", "傳送者", "項目", "金額", "類別", "原始訊息"]:
            converted = [HEADERS]
            for row in old_rows:
                row = row + [""] * max(0, 7 - len(row))
                converted.append([
                    row[0], row[1], row[2], "支出", row[3], row[4], row[5], row[6]
                ])
            ws.clear()
            ws.update("A1", converted)
        else:
            # 不認得的表頭不自動清除，避免意外覆蓋使用者資料。
            raise RuntimeError(
                "「記帳」工作表的第一列欄位與程式不一致。請先備份試算表，再確認第一列欄名。"
            )
    return ws


def parse_date_hint(text):
    """先用明確規則辨識今天、昨天、前天及常見日期格式。"""
    today = now_tw().date()
    cleaned = text

    if "大前天" in cleaned:
        target = today - datetime.timedelta(days=3)
        cleaned = cleaned.replace("大前天", " ")
    elif "前天" in cleaned:
        target = today - datetime.timedelta(days=2)
        cleaned = cleaned.replace("前天", " ")
    elif "昨天" in cleaned or "昨日" in cleaned:
        target = today - datetime.timedelta(days=1)
        cleaned = cleaned.replace("昨天", " ").replace("昨日", " ")
    elif "今天" in cleaned or "今日" in cleaned:
        target = today
        cleaned = cleaned.replace("今天", " ").replace("今日", " ")
    else:
        target = today

    # 支援 2026/10/15、2026-10-15、2026.10.15
    match = re.search(r"(?<!\d)(20\d{2})[/-](\d{1,2})[/-](\d{1,2})(?!\d)", cleaned)
    if match:
        try:
            target = datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            cleaned = cleaned[:match.start()] + " " + cleaned[match.end():]
        except ValueError:
            pass
    else:
        # 支援 10/15（使用今年；若日期看起來是未來日期，則不擅自改成年份）
        match = re.search(r"(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)", cleaned)
        if match:
            try:
                target = datetime.date(today.year, int(match.group(1)), int(match.group(2)))
                cleaned = cleaned[:match.start()] + " " + cleaned[match.end():]
            except ValueError:
                pass

    cleaned = re.sub(r"^\s*(的|那天|當天)+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ，,。:：")
    return target, cleaned


def infer_main_type(text):
    """優先辨識使用者明確輸入的大項目；沒寫時再依常見詞彙推斷。"""
    s = text.strip()
    if re.search(r"^(刷卡|信用卡|卡刷)(?:\s|[:：,，|｜-]|$)", s):
        return "刷卡"
    if re.search(r"^(收入|入帳|進帳)(?:\s|[:：,，|｜-]|$)", s):
        return "收入"
    if re.search(r"^(支出|花費|付款)(?:\s|[:：,，|｜-]|$)", s):
        return "支出"

    # 明確收入項目
    if re.search(r"(薪資|薪水|工資|獎金|利息收入|投資收入|兼職收入|收到紅包|收入|退款入帳)", s):
        return "收入"
    # 「交卡費」是支出：代表實際把款項付出去。
    if "交卡費" in s or "繳卡費" in s or "繳信用卡" in s:
        return "支出"
    # 未指定時，預設支出。
    return "支出"


def clean_main_prefix(text, main_type):
    s = text.strip()
    prefixes = {
        "刷卡": r"^(刷卡|信用卡|卡刷)\s*[:：,，|｜-]?\s*",
        "收入": r"^(收入|入帳|進帳)\s*[:：,，|｜-]?\s*",
        "支出": r"^(支出|花費|付款)\s*[:：,，|｜-]?\s*",
    }
    s = re.sub(prefixes[main_type], "", s)
    return s.strip()


def default_category(item, main_type):
    s = item.lower()
    if main_type == "收入":
        if any(k in s for k in ["薪水", "薪資", "工資", "月薪"]):
            return "薪資"
        if "獎金" in s:
            return "獎金"
        if "紅包" in s:
            return "紅包"
        if "退款" in s or "退費" in s:
            return "退款"
        if "利息" in s:
            return "利息"
        if "投資" in s or "股息" in s or "配息" in s:
            return "投資收入"
        return "其他"

    if "交卡費" in s or "繳卡費" in s or "繳信用卡" in s:
        return "交卡費"
    if any(k in s for k in ["食材", "食物", "食品", "生鮮", "蔬菜", "水果", "肉類",
                             "海鮮", "零食", "點心", "飲料", "早餐", "午餐", "晚餐",
                             "宵夜", "飯", "麵", "鍋", "滷味", "咖啡", "便當", "餐"]):
        return "餐飲"
    if any(k in s for k in ["規費", "手續費", "辦理費", "地政", "戶政", "登記費", "代辦", "公證", "稅", "罰款"]):
        return "稅費"
    if any(k in s for k in ["停車", "加油", "捷運", "公車", "火車", "高鐵", "計程車", "車資", "交通"]):
        return "交通"
    if any(k in s for k in ["醫院", "診所", "藥", "看醫生", "掛號", "復健"]):
        return "醫療健康"
    if any(k in s for k in ["保險"]):
        return "保險"
    if any(k in s for k in ["訂閱", "會員", "月費", "netflix", "spotify", "claude", "chatgpt"]):
        return "訂閱服務"
    return "其他"


def fallback_parse(text):
    """OpenAI 無法使用時的基本備援：抓最後一個金額，避免完全不能記帳。"""
    target_date, cleaned = parse_date_hint(text)
    main_type = infer_main_type(cleaned)
    cleaned = clean_main_prefix(cleaned, main_type)
    match = re.search(r"^(.*?)[\s$＄,，:：|｜-]*(\d[\d,]*)\s*(元|塊|塊錢)?$", cleaned)
    if not match:
        return None
    item = match.group(1).strip(" ，,。:：|｜-")
    amount_text = match.group(2).replace(",", "")
    if not amount_text.isdigit() or not item:
        return None
    amount = int(amount_text)
    if amount < 0:
        return None
    return {
        "date": target_date.isoformat(),
        "main_type": main_type,
        "item": item,
        "amount": amount,
        "category": default_category(item, main_type),
    }


def parse_record(text):
    """用 AI 理解自然語言；日期先由程式規則處理，AI 只協助判斷項目與類別。"""
    target_date, cleaned = parse_date_hint(text)
    main_type = infer_main_type(cleaned)
    cleaned = clean_main_prefix(cleaned, main_type)

    if not client:
        return fallback_parse(text)

    categories = INCOME_CATEGORIES if main_type == "收入" else EXPENSE_CATEGORIES
    prompt = f"""
你是繁體中文記帳助手。請把使用者訊息解析成一筆記帳資料。
使用者原始訊息：{text}
移除日期和大項目提示後的文字：{cleaned}
已由程式判定的大項目：{main_type}
記帳日期（已由程式依台灣時間判定）：{target_date.isoformat()}

只能回傳一個 JSON 物件，不要 Markdown 或其他文字，格式：
{{"item":"項目名稱","amount":100,"category":"餐飲"}}

規則：
- item 保留主要項目名稱，不要把日期詞或「支出／收入／刷卡」當成項目名稱。
- amount 是金額整數。忽略金額中的逗號，例如 1,200 是 1200。
- category 必須從這些類別選一個：{", ".join(categories)}
- 支出與刷卡沿用相同小項目。
- 含食材、食物、食品、生鮮、蔬菜、水果、肉類、海鮮、零食、點心、飲料、早餐、午餐、晚餐、宵夜、飯、麵、鍋、滷味、便當、咖啡等詞，歸「餐飲」。
- 含規費、手續費、辦理費、地政、戶政、登記費、代辦、公證、稅、罰款等詞，歸「稅費」。
- 停車、捷運、公車、加油、車資等交通項目，歸「交通」。
- 「交卡費／繳卡費／繳信用卡」若大項目是支出，類別選「交卡費」。
- 不要把刷卡消費當成支出；如果訊息開頭明確是「刷卡」，大項目已是刷卡。
- 若沒有金額或不是記帳內容，回傳 {{"item":"","amount":0,"category":"其他"}}。
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=250,
            temperature=0
        )
        raw = (response.choices[0].message.content or "").strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        data = json.loads(raw)
        item = str(data.get("item", "")).strip()
        amount = int(data.get("amount", 0))
        category = str(data.get("category", "其他")).strip()
        allowed = INCOME_CATEGORIES if main_type == "收入" else EXPENSE_CATEGORIES
        if not item or amount <= 0 or category not in allowed:
            return fallback_parse(text)
        return {
            "date": target_date.isoformat(),
            "main_type": main_type,
            "item": item,
            "amount": amount,
            "category": category,
        }
    except Exception as e:
        sys.stdout.write(f"[PARSE] OpenAI 解析失敗，改用備援：{e}\n")
        return fallback_parse(text)


def reply_message(reply_token, text):
    if not reply_token or not LINE_CHANNEL_ACCESS_TOKEN:
        sys.stdout.write("[LINE] 缺少 replyToken 或 LINE_CHANNEL_ACCESS_TOKEN，無法回覆\n")
        return
    url = "https://api.line.me/v2/bot/message/reply"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    try:
        response = req.post(
            url,
            headers=headers,
            json={"replyToken": reply_token, "messages": [{"type": "text", "text": text}]},
            timeout=15
        )
        if response.status_code >= 400:
            sys.stdout.write(f"[LINE] 回覆失敗 {response.status_code}: {response.text}\n")
    except Exception as e:
        sys.stdout.write(f"[LINE] 回覆錯誤：{e}\n")


def get_records():
    ws = get_sheet()
    rows = ws.get_all_records()
    result = []
    for row in rows:
        try:
            amount = int(str(row.get("金額", 0)).replace(",", ""))
            date_str = str(row.get("日期", "")).strip()
            main_type = str(row.get("大項目", "支出")).strip() or "支出"
            if main_type not in MAIN_TYPES:
                main_type = "支出"
            result.append({
                "日期": date_str,
                "時間": str(row.get("時間", "")),
                "傳送者": str(row.get("傳送者", "未知")),
                "大項目": main_type,
                "項目": str(row.get("項目", "")),
                "金額": amount,
                "小項目": str(row.get("小項目", "其他")) or "其他",
                "原始訊息": str(row.get("原始訊息", "")),
            })
        except (ValueError, TypeError):
            continue
    return result


def monthly_total(main_type, today=None):
    today = today or now_tw().date()
    ym = today.strftime("%Y-%m")
    return sum(
        r["金額"] for r in get_records()
        if r["大項目"] == main_type and r["日期"].startswith(ym)
    )


def cumulative_label(main_type):
    labels = {"支出": "本月支出累積", "收入": "本月收入累積", "刷卡": "本月刷卡累積"}
    return labels[main_type]


def save_record(sender, original_text, record):
    ws = get_sheet()
    current_time = now_tw().strftime("%H:%M")
    ws.append_row([
        record["date"], current_time, sender, record["main_type"],
        record["item"], record["amount"], record["category"], original_text
    ], value_input_option="USER_ENTERED")


def query_stats(period):
    """今日／本月完整總覽：列出每筆明細、分類加總，以及三大項目各自的合計。"""
    try:
        records = get_records()
    except Exception as e:
        return f"📭 無法連接試算表：{e}"

    today = now_tw().date()
    if period == "today":
        filtered = [r for r in records if r["日期"] == today.isoformat()]
        title = f"📅 {today.strftime('%Y/%m/%d')} 今日記帳總覽"
        empty_text = "今天還沒有記帳資料。"
    else:
        ym = today.strftime("%Y-%m")
        filtered = [r for r in records if r["日期"].startswith(ym)]
        title = f"📊 {today.strftime('%Y年%m月')} 本月記帳總覽"
        empty_text = "本月還沒有記帳資料。"

    if not filtered:
        return f"{title}\n\n{empty_text}\n\n💸 支出：$0\n💵 收入：$0\n💳 刷卡：$0"

    # 固定依照支出、收入、刷卡順序顯示；各大項目獨立計算。
    section_icons = {"支出": "💸", "收入": "💵", "刷卡": "💳"}
    lines = [title, ""]

    for main_type in MAIN_TYPES:
        type_rows = [r for r in filtered if r["大項目"] == main_type]
        total = sum(r["金額"] for r in type_rows)
        lines.append(f"{section_icons[main_type]} {main_type}明細")

        if not type_rows:
            lines.append("  （沒有資料）")
            lines.append(f"  合計：$0")
            lines.append("")
            continue

        # 每筆明細都列出日期（跨月回填或本月查詢時較清楚）、項目、金額與小項目。
        type_rows.sort(key=lambda r: (r["日期"], r["時間"], r["項目"]))
        for r in type_rows:
            date_label = r["日期"][5:] if len(r["日期"]) >= 10 else r["日期"]
            category = r["小項目"] or "其他"
            lines.append(f"  • {date_label} {r['項目']} ${r['金額']:,}（{category}）")

        # 同時顯示小項目分類加總。
        category_totals = {}
        for r in type_rows:
            category = r["小項目"] or "其他"
            category_totals[category] = category_totals.get(category, 0) + r["金額"]

        lines.append("  分類小計：")
        for category, amount in sorted(category_totals.items(), key=lambda x: (-x[1], x[0])):
            lines.append(f"    └ {category}：${amount:,}")
        lines.append(f"  **{main_type}合計：${total:,}**")
        lines.append("")

    # 明確提醒：信用卡消費與實際繳卡費分開列示，不把兩者合併。
    lines.append("ℹ️ 刷卡消費與「支出－交卡費」分開統計，避免混在一起。")
    return "\n".join(lines)


def query_people_stats():
    try:
        records = get_records()
    except Exception as e:
        return f"📭 無法連接試算表：{e}"
    ym = now_tw().date().strftime("%Y-%m")
    filtered = [r for r in records if r["日期"].startswith(ym)]
    if not filtered:
        return "📊 本月各人統計\n\n目前沒有資料。"

    totals = {}
    for r in filtered:
        key = (r["傳送者"], r["大項目"])
        totals[key] = totals.get(key, 0) + r["金額"]

    lines = [f"👥 {now_tw().strftime('%Y年%m月')} 各人統計", ""]
    for (person, main_type), amount in sorted(totals.items(), key=lambda x: (x[0][0], x[0][1])):
        lines.append(f"{person}｜{main_type}：${amount:,}")
    return "\n".join(lines)


def handle_event(event):
    if event.get("type") != "message":
        return
    msg = event.get("message", {})
    if msg.get("type") != "text":
        return

    text = msg.get("text", "").strip()
    reply_token = event.get("replyToken", "")
    sender = event.get("source", {}).get("userId", "未知")
    sys.stdout.write(f"[MSG] {text!r} from {sender}\n")

    commands = {
        "/今日": lambda: query_stats("today"),
        "/today": lambda: query_stats("today"),
        "/本月": lambda: query_stats("month"),
        "/month": lambda: query_stats("month"),
        "/統計": query_people_stats,
        "/stats": query_people_stats,
    }
    if text in commands:
        reply_message(reply_token, commands[text]())
        return

    if text in ["/說明", "/help"]:
        reply_message(
            reply_token,
            "📒 記帳小工具｜使用說明\n\n"
            "直接記帳（預設支出）：\n早餐 100\n昨天的早餐 100\n前天停車費 80\n\n"
            "收入：\n收入 薪資 30000\n\n"
            "刷卡：\n刷卡 午餐 150\n昨天刷卡 交通 120\n\n"
            "繳信用卡費（實際付款，記為支出）：\n交卡費 5000\n\n"
            "日期也可寫：2026/10/15 早餐 100\n\n"
            "查詢：/今日、/本月、/統計、/說明"
        )
        return

    record = parse_record(text)
    if not record:
        return

    try:
        save_record(sender, text, record)
        total = monthly_total(record["main_type"])
        label = cumulative_label(record["main_type"])
        saved_date = datetime.date.fromisoformat(record["date"])
        date_note = ""
        if saved_date != now_tw().date():
            date_note = f"\n📅 記錄日期：{saved_date.strftime('%Y/%m/%d')}"
        response_text = (
            f"✅ {record['main_type']}｜{record['item']} ${record['amount']:,}"
            f"（{record['category']}）{date_note}\n"
            f"📊 {label}：${total:,}"
        )
        reply_message(reply_token, response_text)
        sys.stdout.write(f"[SHEETS] 寫入成功：{record}\n")
    except Exception as e:
        sys.stdout.write(f"[SHEETS] 寫入失敗：{e}\n")
        reply_message(reply_token, f"⚠️ 這筆記帳沒有成功寫入，請稍後再試。\n原因：{str(e)[:250]}")


@app.route("/callback", methods=["POST"])
def callback():
    body = request.get_data(as_text=True)
    if not LINE_CHANNEL_SECRET:
        sys.stdout.write("[WEBHOOK] 沒設定 LINE_CHANNEL_SECRET，拒收\n")
        return "LINE_CHANNEL_SECRET not set", 500

    expected = base64.b64encode(
        hmac.new(LINE_CHANNEL_SECRET.encode(), body.encode(), hashlib.sha256).digest()
    ).decode()
    received = request.headers.get("X-Line-Signature", "")
    if not hmac.compare_digest(expected, received):
        sys.stdout.write("[WEBHOOK] 簽章不對，拒收\n")
        return "bad signature", 403

    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        return "OK"

    for event in data.get("events", []):
        thread = threading.Thread(target=handle_event, args=(event,), daemon=True)
        thread.start()
    return "OK"


@app.route("/", methods=["GET"])
def health():
    return "✅ 記帳小工具運作中"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
