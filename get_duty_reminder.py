import urllib.request
import urllib.parse
import json
import base64
import zlib
import re
import datetime

# ================= 配置区 =================
TENCENT_DOC_URL = "https://docs.qq.com/sheet/DY2Z5dGpBY1p4T0xo"

# ⚠️ 请将下方引号内的文字替换为您从微信测试号获取的真实数据
APPID = "wxa51aa91318272a31"
APPSECRET = "bf5a9a751bbb55e67056252918c7b6c6"
OPENID = "o2kfK26g6hXFgHM14r71WduIlosU"
TEMPLATE_ID = "0sm8dL27YrzvAG607iHs2R-bf42Wf9IQ7ZmsmrFsuM8"
# ==========================================

def get_wechat_access_token():
    """获取微信官方调用凭证"""
    url = f"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={APPID}&secret={APPSECRET}"
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            return data.get("access_token")
    except Exception as e:
        print(f"获取微信 Token 失败: {e}")
        return None

def push_to_wechat_official(date_str, parent_name, phone_num):
    """向微信官方测试号发送模板消息"""
    token = get_wechat_access_token()
    if not token:
        return

    url = f"https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={token}"

    payload = {
        "touser": OPENID,
        "template_id": TEMPLATE_ID,
        "data": {
            "first": {"value": "🔔 各位家长好，今日晚自习值班提醒：\n", "color": "#173177"},
            "date": {"value": date_str, "color": "#173177"},
            "parent": {"value": parent_name, "color": "#ff0000"},
            "phone": {"value": phone_num, "color": "#173177"},
            "remark": {"value": "\n⚠️ 请于 18:00 前到达学校，凭孩子姓名+班级在门卫处登记进校。\n到校后可在群内报备一声。感谢支持！", "color":
"#333333"}
        }
    }

    data_bytes = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"})

    try:
        with urllib.request.urlopen(req) as resp:
            print("微信官方推送结果:", resp.read().decode('utf-8'))
    except Exception as e:
        print("推送微信时发生异常:", e)

def extract_client_vars(html_content):
    matches = re.findall(r"atob\('([^']+)'\)", html_content)
    if not matches: return None
    m = matches[0]
    missing_padding = len(m) % 4
    if missing_padding: m += '=' * (4 - missing_padding)
    try:
        return json.loads(urllib.parse.unquote(base64.b64decode(m).decode('utf-8', errors='ignore')))
    except Exception: return None

def extract_strings_from_related_sheet(compressed_sheet_b64):
    decoded_b64 = base64.b64decode(compressed_sheet_b64)
    decompressed = zlib.decompress(decoded_b64)
    current_str = []
    strings = []
    for idx, b in enumerate(decompressed):
        is_printable = (32 <= b <= 126) or (0xe4 <= b <= 0xe9) or (0x80 <= b <= 0xbf)
        if is_printable: current_str.append(b)
        else:
            if len(current_str) >= 2:
                try:
                    s = bytes(current_str).decode('utf-8')
                    s_clean = re.sub(r'[\s\x00-\x1f]', '', s)
                    if len(s_clean) >= 2: strings.append(s_clean)
                except: pass
            current_str = []

    def fix_mojibake(s):
        try: return s.encode('gbk', errors='ignore').decode('utf-8', errors='ignore')
        except: return s

    decoded_strings = []
    for s in strings:
        fixed = fix_mojibake(s)
        if not fixed.strip() or len(fixed) < len(s) / 2: fixed = s
        decoded_strings.append(fixed)
    return decoded_strings

def main_handler():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    req = urllib.request.Request(TENCENT_DOC_URL, headers=headers)
    try:
        with urllib.request.urlopen(req) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"请求网页失败: {e}"); return

    client_vars = extract_client_vars(html)
    if not client_vars: return

    pad_info = client_vars.get('docInfo', {}).get('padInfo', {})
    pad_id = f"{pad_info.get('domainId', '300000000')}${pad_info.get('padId', '')}"

    tabs = []
    footer_match = re.search(r'id="footerDomStr"[^>]*>([^<]+)</div>', html)
    if footer_match:
        try:
            for t in json.loads(footer_match.group(1)):
                tab_name = t.get('name', '')
                try: tab_name = tab_name.encode('gbk', errors='ignore').decode('utf-8', errors='ignore')
                except: pass
                tabs.append({'id': t.get('id'), 'name': tab_name})
        except: pass

    if not tabs: return

    today = datetime.date.today()
    current_month_str = f"{today.month}月"
    target_tab = tabs[0]
    for t in tabs:
        if current_month_str in t['name']:
            target_tab = t; break

    print(f"今日日期: {today.strftime('%Y-%m-%d')}，工作表: [{target_tab['name']}]")

    params = {"padId": pad_id, "subId": target_tab['id'], "startrow": 0, "endrow": 150, "normal": 1, "outformat": 1}
    api_url_full = f"https://docs.qq.com/dop-api/get/sheet?{urllib.parse.urlencode(params)}"
    api_req = urllib.request.Request(api_url_full, headers={"User-Agent": headers["User-Agent"], "Referer": TENCENT_DOC_URL})

    try:
        with urllib.request.urlopen(api_req) as resp:
            api_data = json.loads(resp.read().decode('utf-8'))
    except Exception as e:
        print(f"调用 dop-api 失败: {e}"); return

    text_list = api_data.get('data', {}).get('initialAttributedText', {}).get('text', [])
    if not text_list: return
    sheet_strings = extract_strings_from_related_sheet(text_list[0].get('related_sheet', ''))

    search_patterns = [f"{today.month}{today.day}", f"{today.day}号"]
    on_duty_parent = "未安排/未登记"
    parent_phone = "暂无联系方式"

    for idx, val in enumerate(sheet_strings):
        if any(p == val.strip() for p in search_patterns):
            for offset in range(1, 5):
                if idx + offset < len(sheet_strings):
                    next_val = sheet_strings[idx + offset].strip()
                    if next_val.isdigit() and len(next_val) == 11:
                        parent_phone = next_val
                    elif any(k in next_val for k in ["号", "家长", "爸爸", "妈妈"]) or len(next_val) >= 2:
                        if next_val not in ["周一", "周二", "周三", "周四", "周五", "周六", "周日", "星期", "日期", "值班时间", "家长手机号码",
"值班时间为"]:
                            on_duty_parent = next_val
            break

    date_formatted = f"{today.strftime('%Y-%m-%d')} (周{['一','二','三','四','五','六','日'][today.weekday()]})"
    print(f"解析完成：{date_formatted} | {on_duty_parent} | {parent_phone}")

    # 执行微信推送
    push_to_wechat_official(date_formatted, on_duty_parent, parent_phone)

if __name__ == "__main__":
    main_handler()
