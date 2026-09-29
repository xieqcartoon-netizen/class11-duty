import time
import requests
import pandas as pd
import datetime
import os

# --- 配置区 ---
TENCENT_DOC_URL = "https://docs.qq.com/sheet/DY2Z5dGpBY1p4T0xo"
DOC_ID = "DY2Z5dGpBY1p4T0xo"
OUTPUT_FILENAME = "temp_duty_sheet.xlsx"

# 从 GitHub Secrets 中安全读取所有凭证和 Cookie
APPID = os.environ.get("WECHAT_APPID")
APPSECRET = os.environ.get("WECHAT_APPSECRET")
OPENID = os.environ.get("WECHAT_OPENID")
TEMPLATE_ID = os.environ.get("WECHAT_TEMPLATE_ID")
TENCENT_COOKIE = os.environ.get("TENCENT_COOKIE")

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
            "remark": {"value": "\n⚠️ 请于 18:00 前到达学校，凭孩子姓名+班级在门卫处登记进校。\n到校后可在群内报备一声。感谢支持！", "color":
"#333333"}
        }
    }

    try:
        resp = requests.post(url, json=payload)
        print("微信推送结果:", resp.json())
    except Exception as e:
        print("推送微信时发生异常:", e)

def export_and_download_excel():
    """使用 Cookie 调用腾讯官方 API 导出并下载 Excel"""
    if not TENCENT_COOKIE:
        print("错误: 未配置 TENCENT_COOKIE 密钥!")
        return False

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Cookie": TENCENT_COOKIE,
        "Referer": "https://docs.qq.com/",
        "Content-Type": "application/json"
    }

    # 步骤 1: 创建导出任务
    export_url = "https://docs.qq.com/v1/export/export_office"
    payload = {
        "docId": DOC_ID,
        "version": 2,
        "exportSource": "client",
        "exportType": 0,
        "switches": {
            "embedFonts": False
        }
    }

    print("正在向腾讯服务器创建 Excel 导出任务...")
    try:
        resp = requests.post(export_url, json=payload, headers=headers)
        res_json = resp.json()
        if res_json.get("ret") != 0:
            print(f"导出失败: {res_json.get('msg')}")
            return False
        operation_id = res_json.get("operationId")
    except Exception as e:
        print(f"请求导出接口失败: {e}")
        return False

    # 步骤 2: 轮询任务进度
    progress_url = f"https://docs.qq.com/v1/export/query_progress?operationId={operation_id}"
    file_url = None
    print("正在等待腾讯云端文件渲染...")
    for _ in range(15):
        time.sleep(2)
        try:
            progress_resp = requests.get(progress_url, headers=headers)
            progress_json = progress_resp.json()
            progress = progress_json.get("progress", 0)
            print(f"  当前导出进度: {progress}%")
            if progress == 100:
                file_url = progress_json.get("file_url")
                break
        except Exception as e:
            print(f"轮询进度发生错误: {e}")
            return False

    if not file_url:
        print("导出超时，未获得下载链接。")
        return False

    # 步骤 3: 下载 Excel 文件
    print("正在下载导出的 Excel 表格...")
    try:
        file_data = requests.get(file_url)
        with open(OUTPUT_FILENAME, "wb") as f:
            f.write(file_data.content)
        print("Excel 表格成功下载并保存到云端运行环境中！")
        return True
    except Exception as e:
        print(f"下载文件失败: {e}")
        return False

def parse_and_find_duty():
    """使用 pandas 精准按列匹配并提取值班家长"""
    if not os.path.exists(OUTPUT_FILENAME):
        return None, None

    try:
        xl = pd.ExcelFile(OUTPUT_FILENAME)
        sheet_name = xl.sheet_names[0]

        # 寻找当前月份的工作表 (例如 9月值班表)
        today = datetime.date.today()
        current_month_str = f"{today.month}月"
        for name in xl.sheet_names:
            if current_month_str in name:
                sheet_name = name
                break

        print(f"匹配并读取当前月份的工作表: [{sheet_name}]")
        # 跳过空表头，通常值班表实际内容在第 3 行或第 4 行开始（跳过前 2 行标题）
        df = pd.read_excel(OUTPUT_FILENAME, sheet_name=sheet_name, header=2)

        # 自动清洗并定位：日期、值班家长、家长手机号码
        date_col = None
        parent_col = None
        phone_col = None

        # 遍历前 4 列，定位字段名
        for col in df.columns[:5]:
            col_str = str(col).strip()
            # 扫描这一列的前几行，查找关键字
            row_samples = df[col].head(4).astype(str).tolist()
            if any("日期" in r or "星期" in r for r in row_samples) or "日期" in col_str:
                if date_col is None: date_col = col
            if any("家长" in r or "值班" in r for r in row_samples) or "家长" in col_str:
                parent_col = col
            if any("手机" in r or "号码" in r or "联系" in r for r in row_samples) or "手机" in col_str:
                phone_col = col

        # 降级容错备用：如果未自动定位成功，使用前 4 列
        if date_col is None: date_col = df.columns[0]
        if parent_col is None: parent_col = df.columns[2]
        if phone_col is None: phone_col = df.columns[3]

        print(f"🎯 精准定位列 -> 日期列: '{date_col}' | 值班家长列: '{parent_col}' | 手机列: '{phone_col}'")

        # 准备今天日期的多格式匹配方案
        today_formatted_options = [
            today.strftime("%Y-%m-%d"),
            today.strftime("%Y/%m/%d"),
            f"{today.month}/{today.day}",
            f"{today.month}月{today.day}日",
            f"{today.month}月{today.day}",
            str(today.day), # 单数字如 28
            f"{today.day}号"
        ]

        on_duty_parent = "未安排/未登记"
        parent_phone = "暂无联系方式"

        for idx, row in df.iterrows():
            raw_date = row[date_col]
            if pd.isna(raw_date):
                continue

            # 处理 pandas 读入标准日期产生的 Timestamp 格式
            if isinstance(raw_date, datetime.datetime) or hasattr(raw_date, 'strftime'):
                row_date_val = raw_date.strftime("%Y-%m-%d")
            else:
                row_date_val = str(raw_date).strip()

            # 精确匹配今天
            matched = False
            for opt in today_formatted_options:
                if opt == row_date_val or (opt in row_date_val and len(opt) > 2):
                    matched = True
                    break

            if matched:
                on_duty_parent = str(row[parent_col]).strip()
                parent_phone = str(row[phone_col]).strip()

                # 去除 NaN 噪音
                if on_duty_parent == "nan" or not on_duty_parent or on_duty_parent.isdigit():
                    on_duty_parent = "未安排/未登记"
                if parent_phone == "nan" or not parent_phone:
                    parent_phone = "暂无联系方式"

                print(f"✅ 成功定位到本日值班排班 (Row {idx+3}): {row_date_val} | 家长: {on_duty_parent} | 电话: {parent_phone}")
                break

        return on_duty_parent, parent_phone

    except Exception as e:
        print(f"解析 Excel 失败: {e}")
        return None, None

def main_handler():
    # 1. 导出并下载 Excel 报表
    success = export_and_download_excel()
    if not success:
        print("未能成功获取 Excel 报表。")
        return

    # 2. 精确解析
    parent_name, phone_num = parse_and_find_duty()
    if parent_name is None:
        return

    today = datetime.date.today()
    date_formatted = f"{today.strftime('%Y-%m-%d')} (周{['一','二','三','四','五','六','日'][today.weekday()]})"

    # 3. 官方安全通道推送
    push_to_wechat_official(date_formatted, parent_name, phone_num)

if __name__ == '__main__':
    main_handler()
