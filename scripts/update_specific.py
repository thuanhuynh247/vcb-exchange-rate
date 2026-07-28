import requests
import pandas as pd
import os
import sys
from datetime import datetime

# Setup paths
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from get_rates import save_to_excel, logger

API_URL = 'https://www.vietcombank.com.vn/api/exchangerates'
OUTPUT_FILE = r'D:\Tygia-Tudong\TyGia_Banking.xlsx'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
}

CURRENCY_NAMES = {
    'USD': 'US DOLLAR', 'EUR': 'EURO', 'GBP': 'POUND STERLING',
    'JPY': 'YEN', 'AUD': 'AUSTRALIAN DOLLAR', 'CAD': 'CANADIAN DOLLAR',
    'CHF': 'SWISS FRANC', 'SGD': 'SINGAPORE DOLLAR', 'CNY': 'YUAN RENMINBI',
    'HKD': 'HONGKONG DOLLAR', 'THB': 'THAILAND BAHT', 'KRW': 'KOREAN WON',
    'MYR': 'MALAYSIAN RINGGIT', 'INR': 'INDIAN RUPEE', 'KWD': 'KUWAITI DINAR',
    'SAR': 'SAUDI RIAL', 'NOK': 'NORWEGIAN KRONER', 'SEK': 'SWEDISH KRONA',
    'DKK': 'DANISH KRONE', 'RUB': 'RUSSIAN RUBLE',
}

def _to_float(val):
    try:
        v = float(str(val).replace(',', ''))
        return v if v > 0 else None
    except (ValueError, TypeError):
        return None

def fetch_rates_for_date(date_str: str) -> list[dict] | None:
    try:
        resp = requests.get(
            API_URL, params={'date': date_str},
            headers=HEADERS, timeout=15
        )
        resp.raise_for_status()
        data = resp.json()

        api_date = data.get('Date', '')[:10]  # 'YYYY-MM-DD'
        rows = []
        for item in data.get('Data', []):
            code = item.get('currencyCode', '')
            if code:
                rows.append({
                    'Ngày Cập Nhật': api_date,
                    'Mã Ngoại Tệ': code,
                    'Tên Ngoại Tệ': CURRENCY_NAMES.get(code, item.get('currencyName', '').strip()),
                    'Mua Tiền Mặt': _to_float(item.get('cash', '')),
                    'Mua Chuyển Khoản': _to_float(item.get('transfer', '')),
                    'Bán': _to_float(item.get('sell', '')),
                })
        return rows if rows else None
    except Exception as e:
        logger.error(f"Lỗi khi lấy dữ liệu ngày {date_str}: {e}")
        return None

if __name__ == '__main__':
    target_date = '2026-05-27'
    logger.info(f"Bắt đầu cập nhật tỷ giá ngày {target_date}...")
    rates = fetch_rates_for_date(target_date)
    if rates:
        logger.info(f"Đã lấy thành công {len(rates)} ngoại tệ từ API.")
        save_to_excel(rates, OUTPUT_FILE)
        logger.info(f"Hoàn tất cập nhật tỷ giá ngày {target_date} vào Excel.")
    else:
        logger.error(f"Không thể lấy tỷ giá ngày {target_date}!")
