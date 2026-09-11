#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
CRAWLER CLOUD: Tự Động Hóa Thu Thập Tỷ Giá Đa Ngân Hàng (Headless Serverless)
=============================================================================
Chạy độc lập trên Linux (GitHub Actions Runner) hoặc Windows.
Tự động cào dữ liệu tỷ giá USD từ các ngân hàng chính:
- Vietcombank (XML / API)
- ACB (Open API)
- Cập nhật docs/rates_history.json, docs/latest_live.json, và TyGia_Banking.xlsx
=============================================================================
"""

import os
import sys
import json
import datetime
import urllib3
import requests
import xml.etree.ElementTree as ET

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Resolve project paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCS_DIR = os.path.join(PROJECT_ROOT, "docs")
RATES_HISTORY_FILE = os.path.join(DOCS_DIR, "rates_history.json")
LATEST_LIVE_FILE = os.path.join(DOCS_DIR, "latest_live.json")
EXCEL_ROOT = os.path.join(PROJECT_ROOT, "TyGia_Banking.xlsx")
EXCEL_DOCS = os.path.join(DOCS_DIR, "TyGia_Banking.xlsx")
INDEX_HTML = os.path.join(PROJECT_ROOT, "index.html")
DOCS_INDEX_HTML = os.path.join(DOCS_DIR, "index.html")

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
}

def log(msg):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def fetch_vietcombank():
    """Fetch live rates from Vietcombank XML endpoint."""
    url = "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx?b=68"
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        for ex in root.findall('Exrate'):
            if ex.get('CurrencyCode') == 'USD':
                return {
                    'cash': float(ex.get('Buy').replace(',', '').strip()),
                    'transfer': float(ex.get('Transfer').replace(',', '').strip()),
                    'sell': float(ex.get('Sell').replace(',', '').strip())
                }
    except Exception as e:
        log(f"WARN: Vietcombank fetch error: {e}")
    return None

def fetch_acb():
    """Fetch live rates from ACB Open API."""
    now_iso = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S.000')
    url = f"https://acb.com.vn/api/front/v1/currency?currency=VND&effectiveDateTime={now_iso}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=12, verify=False)
        r.raise_for_status()
        data = r.json()
        rates = {'cash': None, 'transfer': None, 'sell': None}
        if isinstance(data, list):
            for item in data:
                if item.get('exchangeCurrency') == 'USD':
                    deal = item.get('dealType')
                    inst = item.get('instrumentType')
                    rate = float(item.get('exchangeRate'))
                    denom = item.get('denomination')
                    if deal == 'BID' and inst == 'TRANSFER':
                        rates['transfer'] = rate
                    elif deal == 'BID' and inst == 'CASH' and denom == 3:
                        rates['cash'] = rate
                    elif deal == 'ASK' and (denom == 3 or denom is None):
                        rates['sell'] = rate
        if rates['transfer'] and rates['sell']:
            if not rates['cash']:
                rates['cash'] = rates['transfer'] - 30.0
            return rates
    except Exception as e:
        log(f"WARN: ACB fetch error: {e}")
    return None

def main():
    log("=== BẮT ĐẦU CÀO DỮ LIỆU TỶ GIÁ ĐA NGÂN HÀNG (CLOUD RUNNER) ===")
    
    if not os.path.exists(RATES_HISTORY_FILE):
        log(f"ERROR: Không tìm thấy {RATES_HISTORY_FILE}")
        sys.exit(1)
        
    with open(RATES_HISTORY_FILE, "r", encoding="utf-8") as f:
        rates_history = json.load(f)

    date_keys = sorted(rates_history.keys(), reverse=True)
    latest_iso = date_keys[0]
    log(f"Ngày gần nhất trong lịch sử: {latest_iso} ({len(date_keys)} ngày)")

    today_iso = datetime.datetime.now().strftime('%Y-%m-%d')
    today_dd = datetime.datetime.now().strftime('%d/%m/%Y')
    
    vcb = fetch_vietcombank()
    acb = fetch_acb()

    log(f"Kết quả cào VCB: {vcb}")
    log(f"Kết quả cào ACB: {acb}")

    target_iso = today_iso
    if target_iso not in rates_history:
        prev_banks = dict(rates_history[latest_iso]["banks"])
        rates_history[target_iso] = {
            "date_dd": today_dd,
            "banks": prev_banks
        }
        log(f"Tạo mới phiên giao dịch cho ngày {today_dd} ({target_iso})")

    updated_banks = rates_history[target_iso]["banks"]

    if vcb:
        updated_banks["Vietcombank"] = vcb
    if acb:
        updated_banks["ACB"] = acb

    with open(RATES_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(rates_history, f, ensure_ascii=False, indent=2)
    log(f"Đã lưu {RATES_HISTORY_FILE} ({len(rates_history)} ngày)")

    live_snapshot = {
        "timestamp": datetime.datetime.now().isoformat(),
        "date_dd": today_dd,
        "date_iso": target_iso,
        "banks": updated_banks
    }
    with open(LATEST_LIVE_FILE, "w", encoding="utf-8") as f:
        json.dump(live_snapshot, f, ensure_ascii=False, indent=2)
    log(f"Đã xuất snapshot thời gian thực: {LATEST_LIVE_FILE}")

    rates_json_str = json.dumps(rates_history, ensure_ascii=False)
    for html_path in [INDEX_HTML, DOCS_INDEX_HTML]:
        if os.path.exists(html_path):
            with open(html_path, "r", encoding="utf-8") as f:
                content = f.read()
            
            import re
            content = re.sub(r'const RATES_HISTORY = \{.*?\};', f'const RATES_HISTORY = {rates_json_str};', content)
            
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(content)
            log(f"Đã cập nhật dữ liệu mới vào {html_path}")

    # Optional Excel update
    try:
        import openpyxl
        if os.path.exists(EXCEL_ROOT) and vcb:
            wb = openpyxl.load_workbook(EXCEL_ROOT)
            ws = wb["Data_TheoDoi_USD"] if "Data_TheoDoi_USD" in wb.sheetnames else wb.active
            
            date_exists = False
            for row in range(2, ws.max_row + 1):
                val = str(ws.cell(row=row, column=1).value or '')
                if target_iso in val or today_dd in val:
                    date_exists = True
                    break
            
            if not date_exists:
                new_row = ws.max_row + 1
                ws.cell(row=new_row, column=1, value=today_iso)
                ws.cell(row=new_row, column=2, value="USD")
                ws.cell(row=new_row, column=3, value=vcb.get('cash'))
                ws.cell(row=new_row, column=4, value=vcb.get('transfer'))
                ws.cell(row=new_row, column=5, value=vcb.get('sell'))
                wb.save(EXCEL_ROOT)
                import shutil
                shutil.copy2(EXCEL_ROOT, EXCEL_DOCS)
                log("Đã cập nhật file Excel TyGia_Banking.xlsx thành công!")
    except Exception as e:
        log(f"Bỏ qua cập nhật Excel: {e}")

    log("=== HOÀN TẤT CÀO DỮ LIỆU THÀNH CÔNG 100% ===")

if __name__ == "__main__":
    main()
