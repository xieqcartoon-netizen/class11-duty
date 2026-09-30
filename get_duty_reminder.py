import time
import requests
import datetime
import os
import re
import urllib.parse
import json
import base64
import zlib
import struct

# --- 配置区 ---
TENCENT_DOC_URL = "https://docs.qq.com/sheet/DY2Z5dGpBY1p4T0xo"

# ⚠️ 请将下方引号内的文字替换为您从微信测试号获取的真实数据
APPID = "wxa51aa91318272a31"
APPSECRET = "bf5a9a751bbb55e67056252918c7b6c6"
OPENID = "o2kfK26g6hXFgHM14r71WduIlosU"
TEMPLATE_ID = "0sm8dL27YrzvAG607iHs2R-bf42Wf9IQ7ZmsmrFsuM8"
# ==========================================

def get_wechat_access_token():
    url = f"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={APPID}&secret={APPSECRET}"
    try:
        resp = requests.get(url)
        return resp.json().get("access_token")
    except Exception as e:
        print(f"获取微信 Token 失败: {e}")
        return None

def push_to_wechat_official(date_str, parent_name, phone_num):
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
            "remark": {
                "value": """ ⚠️ 请于 21:00 前到达学校，凭孩子姓名+班级在门卫处登记进校。 到校后可在群内报备一声。感谢支持！""",
                "color": "#333333"
            }
        }
    }

    try:
        resp = requests.post(url, json=payload)
        print("微信推送结果:", resp.json())
    except Exception as e:
        print("推送微信时发生异常:", e)

def extract_client_vars(html_content):
    # 自愈搜索：匹配出包含 docInfo 和 padData 的正确 atob 加密块
    matches = re.findall(r"atob\('([^']+)'\)", html_content)
    if not matches:
        print("未在网页源码中搜索到 atob('...') 数据块。")
        return None

    for idx, m in enumerate(matches):
        missing_padding = len(m) % 4
        if missing_padding:
            m += '=' * (4 - missing_padding)
        try:
            decoded_b64 = base64.b64decode(m).decode('utf-8', errors='ignore')
            unquoted = urllib.parse.unquote(decoded_b64)
            data = json.loads(unquoted)
            if 'docInfo' in data or 'padData' in data:
                return data
        except Exception:
            pass
    return None

def extract_strings_from_related_sheet(compressed_sheet_b64):
    decoded_b64 = base64.b64decode(compressed_sheet_b64)
    decompressed = zlib.decompress(decoded_b64)

    current_str = []
    strings = []
    for idx, b in enumerate(decompressed):
        is_printable = (32 <= b <= 126) or (0xe4 <= b <= 0xe9) or (0x80 <= b <= 0xbf)
        if is_printable:
            current_str.append(b)
        else:
            if len(current_str) >= 2:
                try:
                    s = bytes(current_str).decode('utf-8')
                    s_clean = re.sub(r'[\s\x00-\x1f]', '', s)
                    if len(s_clean) >= 2:
                        strings.append((idx - len(current_str), s_clean))
                except:
                    pass
            current_str = []
    return strings, decompressed

def fix_mojibake(s):
    try:
        return s.encode('gbk', errors='ignore').decode('utf-8', errors='ignore')
    except:
        return s

def clean_parent_name(s):
    """自适应还原 GBK 乱码，同时保护已经正常的 UTF-8 中文不被二次破坏"""
    try:
        fixed = s.encode('gbk', errors='ignore').decode('utf-8', errors='ignore')
        raw_chinese = len(re.findall(r'[\u4e00-\u9fff]', s))
        fixed_chinese = len(re.findall(r'[\u4e00-\u9fff]', fixed))
        if fixed_chinese >= raw_chinese - 1 and fixed_chinese > 0:
            return fixed
    except Exception:
        pass
    return s

def main_handler():
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://docs.qq.com/"
    }

    print("1. 正在抓取腾讯文档公开主页并解析内部 padId...")
    try:
        resp = requests.get(TENCENT_DOC_URL, headers=headers)
        if resp.status_code != 200:
            print("抓取网页失败！")
            return
        html = resp.text
    except Exception as e:
        print(f"请求网页失败: {e}"); return

    client_vars = extract_client_vars(html)
    if not client_vars:
        print("解析 clientVars 失败，请确认文档仍处于公开分享状态。")
        return

    pad_info = client_vars.get('docInfo', {}).get('padInfo', {})
    domain_id = pad_info.get('domainId', '300000000')
    pad_id_short = pad_info.get('padId', '')
    pad_id = f"{domain_id}${pad_id_short}"

    # 提取 Tabs 标签列表
    tabs = []
    footer_match = re.search(r'id="footerDomStr"[^>]*>([^<]+)</div>', html)
    if footer_match:
        try:
            for t in json.loads(footer_match.group(1)):
                tab_name = t.get('name', '')
                try:
                    tab_name = tab_name.encode('gbk', errors='ignore').decode('utf-8', errors='ignore')
                except:
                    pass
                tabs.append({'id': t.get('id'), 'name': tab_name})
        except:
            pass

    if not tabs:
        print("未获取到子标签页。")
        return

    # 定位当前月份工作表
    today = datetime.date.today()
    current_month_str = f"{today.month}月"
    target_tab = tabs[0]
    for t in tabs:
        if current_month_str in t['name']:
            target_tab = t
            break

    print(f"2. 匹配当前工作表: [{target_tab['name']}]  (ID: {target_tab['id']})")

    # 3. 直接调用 unauthenticated dop-api 获取第一手原始单元格流
    print("3. 调用 dop-api 拉取单元格数据流...")
    params = {
        "padId": pad_id,
        "subId": target_tab['id'],
        "startrow": 0,
        "endrow": 150,
        "normal": 1,
        "outformat": 1
    }
    api_url = f"https://docs.qq.com/dop-api/get/sheet?{urllib.parse.urlencode(params)}"

    try:
        api_resp = requests.get(api_url, headers={"User-Agent": headers["User-Agent"], "Referer": TENCENT_DOC_URL})
        api_data = api_resp.json()
    except Exception as e:
        print(f"调用 dop-api 失败: {e}")
        return

    text_list = api_data.get('data', {}).get('initialAttributedText', {}).get('text', [])
    if not text_list:
        print("中继数据文本列表为空。")
        return

    related_sheet_compressed = text_list[0].get('related_sheet', '')
    if not related_sheet_compressed:
        print("中继数据中无压缩工作表。")
        return

    # 4. 解析二进制数据流
    print("4. 解析并解压中继数据流...")
    strings, decompressed = extract_strings_from_related_sheet(related_sheet_compressed)

    # 提取日期序列：9月1日到30日对应的 Excel 双精度浮点序列 (46267.0 到 46296.0)
    # 根据当前年份和月份自适应计算起止的 Excel 序列号
    # Excel 1900 日期系统中 2026-09-01 是 46267
    base_excel_date = datetime.date(1899, 12, 30)
    month_start_date = datetime.date(today.year, today.month, 1)
    if today.month == 12:
        month_end_date = datetime.date(today.year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        month_end_date = datetime.date(today.year, today.month + 1, 1) - datetime.timedelta(days=1)

    start_excel_val = (month_start_date - base_excel_date).days
    end_excel_val = (month_end_date - base_excel_date).days

    # 获取此工作表中所有的有效日期浮点数序列
    date_offsets = []
    for d_val in range(start_excel_val, end_excel_val + 1):
        double_bytes = struct.pack("<d", float(d_val))
        idx = decompressed.find(double_bytes)
        if idx != -1:
            date_offsets.append((idx, d_val))
    date_offsets.sort()

    # 提取所有的有效值班家长单元格字符串，排除指令、备注和样式噪音
    parent_names = []
    for offset, raw_s in strings:
        # 正则：必须是带有学号开头的字符串 (例如 "16钮"、"34张鸿毅" 等)
        if re.search(r'^\d+[\u4e00-\u9fff]+', raw_s):
            fixed_s = clean_parent_name(raw_s)
            if not any(k in fixed_s for k in ["值", "表", "校", "期", "提示", "时间", "流程", "备注"]):
                parent_names.append((offset, fixed_s))
    parent_names.sort()

    # 提取所有的 11 位手机号
    phone_numbers = []
    for offset, s in strings:
        if s.isdigit() and len(s) == 11:
            phone_numbers.append((offset, s))
    phone_numbers.sort()

    # 5. 精准进行 1对1 序列化对齐
    today_excel_val = (today - base_excel_date).days

    today_rank_idx = -1
    for idx, (offset, d_val) in enumerate(date_offsets):
        if d_val == today_excel_val:
            today_rank_idx = idx
            break

    if today_rank_idx == -1:
        print(f"本日 ({today.strftime('%m月%d日')}) 没有在排班表的日期列中登记值班，跳过推送。")
        return

    if today_rank_idx >= len(parent_names):
        print(f"排班对齐错误：本日排序为第 {today_rank_idx+1} 个日期，但仅提取出 {len(parent_names)} 个值班家长。")
        return

    on_duty_parent = parent_names[today_rank_idx][1]
    parent_offset = parent_names[today_rank_idx][0]

    # 寻找离当前值班家长最近的电话号码 (通常就在该名字后面 250 字节以内)
    best_phone = "暂无联系方式"
    min_dist = 99999
    for p_offset, p_num in phone_numbers:
        dist = abs(p_offset - parent_offset)
        if dist < min_dist and dist < 250:
            min_dist = dist
            best_phone = p_num

    date_formatted = f"{today.strftime('%Y-%m-%d')} (周{['一','二','三','四','五','六','日'][today.weekday()]})"
    print(f"\n🎉 完美对齐提取结果：")
    print(f"  [日期] {date_formatted}")
    print(f"  [家长] {on_duty_parent}")
    print(f"  [电话] {best_phone}")

    # 6. 安全推送至您的微信测试号
    push_to_wechat_official(date_formatted, on_duty_parent, best_phone)

if __name__ == '__main__':
    main_handler()
