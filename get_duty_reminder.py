import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import datetime
import os
import requests
import re

# --- 配置区 ---
TENCENT_DOC_URL = "https://docs.qq.com/sheet/DY2Z5dGpBY1p4T0xo"
OUTPUT_FILENAME = "temp_duty_sheet.xlsx"

# ⚠️ 请将下方引号内的文字替换为您从微信测试号获取的真实数据
APPID = "wxa51aa91318272a31"
APPSECRET = "bf5a9a751bbb55e67056252918c7b6c6"
OPENID = "o2kfK26g6hXFgHM14r71WduIlosU"
TEMPLATE_ID = "0sm8dL27YrzvAG607iHs2R-bf42Wf9IQ7ZmsmrFsuM8"
TENCENT_COOKIE = os.environ.get("TENCENT_COOKIE") # 从 GitHub Secrets 中读取
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

async def download_via_playwright():
    """使用 Playwright 注入 Cookie 并模拟点击菜单下载 Excel"""
    if not TENCENT_COOKIE:
        print("错误: 未在 GitHub Secrets 中配置 TENCENT_COOKIE 密钥!")
        return False

    async with async_playwright() as p:
        print("1. 正在启动 Headless Chromium 浏览器...")
        browser = await p.chromium.launch(headless=True)

        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="zh-CN"
        )

        # 解析 Cookie 并安全注入
        cookies = []
        for item in TENCENT_COOKIE.split(";"):
            item = item.strip()
            if "=" in item:
                name, value = item.split("=", 1)
                cookies.append({
                    "name": name,
                    "value": value,
                    "domain": ".qq.com",
                    "path": "/"
                })
        await context.add_cookies(cookies)
        print(f"  已成功注入 {len(cookies)} 个腾讯会话 Cookie 认证密钥。")

        page = await context.new_page()
        print("2. 正在加载腾讯文档页面...")
        await page.goto(TENCENT_DOC_URL, wait_until="domcontentloaded")

        # 打印当前实际跳转 of URL
        print(f"  当前浏览器 URL: {page.url}")

        print("  正在等待腾讯云端渲染 Canvas 画布...")
        await page.wait_for_timeout(10000)

        print("3. 正在定位顶部 '文件(File)' 菜单并展开...")
        # 🌟 关键修复：基于您提供的诊断日志，腾讯文档前端团队把“文件”文字藏在了一个内层无文字的 div 中，
        # 并赋予了它 ID "main-menu-file"。我们直接精准打击这个 ID，100% 命中！
        file_menu = page.locator("#main-menu-file").first

        if await file_menu.count() > 0:
            print("  找到了 '文件' 菜单 (ID: main-menu-file)，正在点击...")
            await file_menu.click()
            await page.wait_for_timeout(2000)

             print("4. 正在定位二级菜单 '导出为' 点击展开...")
             export_menu = page.locator("text=导出为").first
             if await export_menu.count() > 0:
                 print("  找到了 '导出为'，正在点击展开...")
                 await export_menu.click() # 🌟 hover() 改为更稳定的 click() 强制展开子菜单
                 await page.wait_for_timeout(2000)
    
                 print("5. 正在定位 '本地 Excel 表格' 下载选项...")
                 # 🌟 精准匹配具有 .dui-menu-item 样式且包含 'Excel' 文字的按钮
                 excel_option = page.locator(".dui-menu-item:has-text('Excel')").first
                 if await excel_option.count() == 0:
                     excel_option = page.locator(".dui-menu-item:has-text('xlsx')").first
                 if await excel_option.count() == 0:
                     excel_option = page.locator("text=本地 Excel").first
    
                 if await excel_option.count() > 0:
                     print("  [触发下载] 正在生成并下载 Excel 流...")
                     try:
                         async with page.expect_download(timeout=25000) as download_info:
                             await excel_option.click()
                         download = await download_info.value
                         await download.save_as(OUTPUT_FILENAME)
                         print(f"🎉 成功! 精准数据已完美下载到云端环境: {OUTPUT_FILENAME}")
                         await browser.close()
                         return True
                     except Exception as ex:
                         print(f"点击下载或保存文件时发生异常: {ex}")
                 else:
                     print("未找到 '本地 Excel' 选项。")
            else:
                print("未找到 '导出为' 菜单选项。")
        else:
            print("未找到 '文件' 菜单按钮！启动自动深度诊断程序...")
            try:
                buttons = await page.locator("button, [role='button'], .menu_menu-button__1Kokg, .btn").all()
                for idx, btn in enumerate(buttons[:10]):
                    b_id = await btn.evaluate("el => el.id")
                    print(f"      [{idx:02d}] ID: '{b_id}'")
            except Exception as e:
                pass
            await page.screenshot(path="error_screenshot.png")
            print("  已截取报错瞬间的浏览器物理画面，并保存为: error_screenshot.png")

        await browser.close()
        return False

def parse_and_find_duty():
    """使用 pandas 精准解析导出的 Excel 并匹配值班家长"""
    if not os.path.exists(OUTPUT_FILENAME):
        print("错误: 未找到本地下载的 Excel 表格文件!")
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

        # 遍历前 5 列，定位列字段
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

            # 处理 pandas 读入 standard 日期产生的 Timestamp 格式
            if isinstance(raw_date, datetime.datetime) or hasattr(raw_date, 'strftime'):
                row_date_val = raw_date.strftime("%Y-%m-%d")
            elif isinstance(raw_date, float):
                # Excel 浮点型日期转换
                try:
                    row_date_val = pd.to_datetime(raw_date, unit='D', origin='1899-12-30').strftime("%Y-%m-%d")
                except:
                    row_date_val = str(int(raw_date))
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
                elif "." in parent_phone:
                    parent_phone = parent_phone.split(".")[0] # 去除浮点数 .0

                print(f"✅ 成功定位到本日值班排班 (Row {idx+4}): {row_date_val} | 家长: {on_duty_parent} | 电话: {parent_phone}")
                break

        return on_duty_parent, parent_phone

    except Exception as e:
        print(f"解析 Excel 失败: {e}")
        return None, None

def main_handler():
    # 步骤 1: 启动浏览器下载文件
    success = asyncio.run(download_via_playwright())
    if not success:
        print("未能成功导出并获取最新的 Excel 报表。")
        return

    # 步骤 2: 解析 Excel 获取值班家长
    parent_name, phone_num = parse_and_find_duty()
    if not parent_name:
        return

    today = datetime.date.today()
    date_formatted = f"{today.strftime('%Y-%m-%d')} (周{['一','二','三','四','五','六','日'][today.weekday()]})"

    # 步骤 3: 官方安全通道推送
    push_to_wechat_official(date_formatted, parent_name, phone_num)

if __name__ == '__main__':
    main_handler()
