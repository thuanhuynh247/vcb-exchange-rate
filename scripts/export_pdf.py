"""
Export PDF Report - Tỷ Giá Ngoại Tệ Vietcombank

Đọc dữ liệu từ TyGia_Banking.xlsx và tạo PDF report cho ngày chọn.

Usage:
    python export_pdf.py                     # Ngày mới nhất
    python export_pdf.py --date 10/03/2026   # Ngày cụ thể (dd/mm/yyyy)
    python export_pdf.py --output report.pdf # Chọn file output
"""
import argparse
import os
import sys
from datetime import datetime

import pandas as pd
from fpdf import FPDF

INPUT_FILE = r'D:\Tygia-Tudong\TyGia_Banking.xlsx'
OUTPUT_DIR = r'D:\Tygia-Tudong'

# Thứ tự hiển thị ngoại tệ chính (giống website VCB)
PRIORITY_CURRENCIES = ['USD', 'EUR', 'GBP', 'JPY', 'AUD', 'CAD', 'CHF', 'SGD',
                       'CNY', 'HKD', 'THB', 'KRW', 'MYR', 'INR', 'KWD',
                       'SAR', 'NOK', 'SEK', 'DKK', 'RUB']


class VCBReport(FPDF):
    """Custom FPDF class for VCB exchange rate report."""

    VCB_GREEN = (0, 112, 60)      # #00703C
    VCB_DARK = (51, 51, 51)       # #333333
    VCB_GRAY = (102, 102, 102)    # #666666
    VCB_WHITE = (255, 255, 255)
    VCB_GREEN_LIGHT = (240, 248, 240)
    VCB_RED = (192, 57, 43)       # #C0392B
    GREEN_UP = (39, 174, 96)      # #27AE60
    BLUE_INFO = (21, 101, 192)    # #1565C0

    def __init__(self, report_date, compare_date=None):
        super().__init__('L', 'mm', 'A4')  # Landscape
        self.report_date = report_date
        self.compare_date = compare_date
        self.set_auto_page_break(auto=True, margin=15)

        # Load DejaVu font (supports Vietnamese Unicode)
        font_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
        font_path = os.path.join(font_dir, 'DejaVuSans.ttf')
        font_bold_path = os.path.join(font_dir, 'DejaVuSans-Bold.ttf')

        if os.path.exists(font_path):
            self.add_font('DejaVu', '', font_path, uni=True)
            self.add_font('DejaVu', 'B', font_bold_path, uni=True)
            self.default_font = 'DejaVu'
        else:
            # Fallback: try to use built-in font (limited Unicode)
            self.default_font = 'Helvetica'

    def header(self):
        # VCB Green header bar
        self.set_fill_color(*self.VCB_GREEN)
        self.rect(0, 0, self.w, 22, 'F')

        # Title
        self.set_font(self.default_font, 'B', 16)
        self.set_text_color(*self.VCB_WHITE)
        self.set_y(4)
        self.cell(0, 14, 'VIETCOMBANK - TY GIA NGOAI TE', align='C')

        # Report date info
        self.set_y(24)
        self.set_font(self.default_font, '', 10)
        self.set_text_color(*self.VCB_GRAY)

        date_str = self.report_date
        report_time = datetime.now().strftime('%H:%M %d/%m/%Y')
        info_text = f'Ngay: {date_str}  |  Xuat bao cao: {report_time}'
        if self.compare_date:
            info_text += f'  |  So sanh voi: {self.compare_date}'
        self.cell(0, 8, info_text, align='C')
        self.ln(12)

    def footer(self):
        self.set_y(-12)
        self.set_font(self.default_font, '', 8)
        self.set_text_color(*self.VCB_GRAY)
        self.cell(0, 10, f'Nguon: Ngan hang TMCP Ngoai thuong Viet Nam (Vietcombank)  |  Trang {self.page_no()}/{{nb}}', align='C')

    def add_rate_table(self, df_today, df_compare=None):
        """Add exchange rate table to PDF."""
        self.set_y(38)

        # Table configuration
        if df_compare is not None:
            col_widths = [12, 32, 40, 40, 35, 35, 25, 35, 25]
            headers = ['STT', 'Ma NT', 'Ten Ngoai Te', 'Mua TM', 'Mua CK',
                       'Ban', '% DoiTD', 'Ban (SS)', '% SS']
        else:
            col_widths = [15, 40, 60, 50, 50, 45]
            headers = ['STT', 'Ma NT', 'Ten Ngoai Te', 'Mua TM', 'Mua CK', 'Ban']

        row_height = 8

        # Header row
        self.set_fill_color(*self.VCB_GREEN)
        self.set_text_color(*self.VCB_WHITE)
        self.set_font(self.default_font, 'B', 9)

        for i, (w, h) in enumerate(zip(col_widths, headers)):
            self.cell(w, row_height + 2, h, border=1, align='C', fill=True)
        self.ln()

        # Data rows
        self.set_font(self.default_font, '', 8)

        def currency_sort_key(code):
            try:
                return PRIORITY_CURRENCIES.index(code)
            except ValueError:
                return 999

        currencies = sorted(df_today['Mã Ngoại Tệ'].unique(), key=currency_sort_key)

        for idx, curr in enumerate(currencies, 1):
            row = df_today[df_today['Mã Ngoại Tệ'] == curr].iloc[0]

            # Alternating row color
            if idx % 2 == 0:
                self.set_fill_color(*self.VCB_GREEN_LIGHT)
                fill = True
            else:
                self.set_fill_color(*self.VCB_WHITE)
                fill = True

            # Cell values
            name = str(row.get('Tên Ngoại Tệ', ''))[:20]
            buy_tm = self._fmt_num(row.get('Mua Tiền Mặt'))
            buy_ck = self._fmt_num(row.get('Mua Chuyển Khoản'))
            sell = self._fmt_num(row.get('Bán'))

            self.set_text_color(*self.VCB_DARK)
            self.cell(col_widths[0], row_height, str(idx), border=1, align='C', fill=fill)

            # Bold currency code
            self.set_font(self.default_font, 'B', 8)
            self.cell(col_widths[1], row_height, curr, border=1, align='C', fill=fill)
            self.set_font(self.default_font, '', 8)

            self.cell(col_widths[2], row_height, name, border=1, align='L', fill=fill)
            self.cell(col_widths[3], row_height, buy_tm, border=1, align='R', fill=fill)
            self.cell(col_widths[4], row_height, buy_ck, border=1, align='R', fill=fill)

            # Sell price in red
            self.set_text_color(*self.VCB_RED)
            self.set_font(self.default_font, 'B', 8)
            self.cell(col_widths[5], row_height, sell, border=1, align='R', fill=fill)
            self.set_font(self.default_font, '', 8)
            self.set_text_color(*self.VCB_DARK)

            # Comparison columns
            if df_compare is not None:
                comp_row = df_compare[df_compare['Mã Ngoại Tệ'] == curr]
                if not comp_row.empty:
                    comp_sell = comp_row.iloc[0].get('Bán')
                    sell_val = row.get('Bán')

                    # % Change vs compare
                    if pd.notna(sell_val) and pd.notna(comp_sell) and comp_sell != 0:
                        pct = (sell_val - comp_sell) / comp_sell * 100
                        pct_str = f'{pct:+.2f}%'
                        if pct > 0:
                            self.set_text_color(*self.GREEN_UP)
                        elif pct < 0:
                            self.set_text_color(*self.VCB_RED)
                    else:
                        pct_str = '-'

                    self.set_font(self.default_font, 'B', 8)
                    self.cell(col_widths[6], row_height, pct_str, border=1, align='C', fill=fill)
                    self.set_font(self.default_font, '', 8)
                    self.set_text_color(*self.BLUE_INFO)
                    self.cell(col_widths[7], row_height, self._fmt_num(comp_sell), border=1, align='R', fill=fill)
                    self.set_text_color(*self.VCB_DARK)

                    # % comparison
                    if pd.notna(sell_val) and pd.notna(comp_sell) and comp_sell != 0:
                        pct2 = (sell_val - comp_sell) / comp_sell * 100
                        pct2_str = f'{pct2:+.2f}%'
                        if abs(pct2) > 0.5:
                            self.set_font(self.default_font, 'B', 9)
                        if pct2 > 0:
                            self.set_text_color(*self.GREEN_UP)
                        elif pct2 < 0:
                            self.set_text_color(*self.VCB_RED)
                    else:
                        pct2_str = '-'

                    self.cell(col_widths[8], row_height, pct2_str, border=1, align='C', fill=fill)
                    self.set_font(self.default_font, '', 8)
                    self.set_text_color(*self.VCB_DARK)
                else:
                    self.cell(col_widths[6], row_height, '-', border=1, align='C', fill=fill)
                    self.cell(col_widths[7], row_height, '-', border=1, align='C', fill=fill)
                    self.cell(col_widths[8], row_height, '-', border=1, align='C', fill=fill)

            self.ln()

    def _fmt_num(self, val):
        """Format number with comma separator."""
        if pd.isna(val) or val is None:
            return '-'
        try:
            v = float(val)
            if v == int(v) and v > 100:
                return f'{int(v):,}'
            return f'{v:,.2f}'
        except (ValueError, TypeError):
            return '-'


def load_data(target_date=None, compare_date=None):
    """Load data from Excel file."""
    if not os.path.exists(INPUT_FILE):
        print(f'❌ File Excel không tồn tại: {INPUT_FILE}')
        sys.exit(1)

    df = pd.read_excel(INPUT_FILE, sheet_name='Data')
    df['_date_key'] = pd.to_datetime(df['Ngày Cập Nhật'], format='mixed').dt.strftime('%Y-%m-%d')

    unique_dates = sorted(df['_date_key'].unique(), reverse=True)
    if not unique_dates:
        print('❌ File Excel không có dữ liệu!')
        sys.exit(1)

    # Resolve target date
    if target_date:
        try:
            dt = datetime.strptime(target_date, '%d/%m/%Y')
            date_key = dt.strftime('%Y-%m-%d')
        except ValueError:
            print(f'❌ Định dạng ngày không hợp lệ: {target_date} (cần dd/mm/yyyy)')
            sys.exit(1)

        if date_key not in unique_dates:
            print(f'❌ Không có dữ liệu cho ngày {target_date}')
            print(f'   Ngày có sẵn gần nhất: {unique_dates[0]}')
            sys.exit(1)
    else:
        date_key = unique_dates[0]

    display_date = datetime.strptime(date_key, '%Y-%m-%d').strftime('%d/%m/%Y')
    df_target = df[df['_date_key'] == date_key]

    # Resolve compare date
    df_comp = None
    comp_display = None
    if compare_date:
        try:
            dt_c = datetime.strptime(compare_date, '%d/%m/%Y')
            comp_key = dt_c.strftime('%Y-%m-%d')
        except ValueError:
            print(f'⚠ Định dạng ngày so sánh không hợp lệ: {compare_date}')
            comp_key = None
        if comp_key and comp_key in unique_dates:
            df_comp = df[df['_date_key'] == comp_key]
            comp_display = datetime.strptime(comp_key, '%Y-%m-%d').strftime('%d/%m/%Y')
    else:
        # Default: previous day
        idx = unique_dates.index(date_key)
        if idx + 1 < len(unique_dates):
            prev_key = unique_dates[idx + 1]
            df_comp = df[df['_date_key'] == prev_key]
            comp_display = datetime.strptime(prev_key, '%Y-%m-%d').strftime('%d/%m/%Y')

    return df_target, display_date, df_comp, comp_display


def main():
    parser = argparse.ArgumentParser(description='Export PDF tỷ giá VCB')
    parser.add_argument('--date', '-d', help='Ngày cần xuất (dd/mm/yyyy). Mặc định: ngày mới nhất')
    parser.add_argument('--compare', '-c', help='Ngày so sánh (dd/mm/yyyy). Mặc định: ngày trước đó')
    parser.add_argument('--output', '-o', help='File PDF output. Mặc định: TyGia_Banking_YYYYMMDD.pdf')
    args = parser.parse_args()

    print('📄 Export PDF tỷ giá VCB...')
    df_today, display_date, df_comp, comp_display = load_data(args.date, args.compare)

    # Output filename
    date_key = datetime.strptime(display_date, '%d/%m/%Y').strftime('%Y%m%d')
    if args.output:
        output_file = args.output
    else:
        output_file = os.path.join(OUTPUT_DIR, f'TyGia_Banking_{date_key}.pdf')

    # Ensure fonts directory exists and download if needed
    _ensure_fonts()

    # Generate PDF
    pdf = VCBReport(display_date, comp_display)
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.add_rate_table(df_today, df_comp)

    pdf.output(output_file)
    print(f'✅ Đã xuất PDF: {output_file}')
    print(f'   📅 Ngày: {display_date}')
    if comp_display:
        print(f'   📊 So sánh với: {comp_display}')
    print(f'   💱 {len(df_today["Mã Ngoại Tệ"].unique())} ngoại tệ')


def _ensure_fonts():
    """Download DejaVu fonts if not present."""
    font_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts')
    os.makedirs(font_dir, exist_ok=True)

    font_files = {
        'DejaVuSans.ttf': 'https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans.ttf',
        'DejaVuSans-Bold.ttf': 'https://github.com/dejavu-fonts/dejavu-fonts/raw/master/ttf/DejaVuSans-Bold.ttf',
    }

    for fname, url in font_files.items():
        fpath = os.path.join(font_dir, fname)
        if not os.path.exists(fpath):
            print(f'📥 Tải font {fname}...')
            try:
                import requests
                resp = requests.get(url, timeout=30)
                resp.raise_for_status()
                with open(fpath, 'wb') as f:
                    f.write(resp.content)
                print(f'   ✅ Đã tải: {fname}')
            except Exception as e:
                print(f'   ⚠ Không thể tải font {fname}: {e}')
                print(f'     PDF sẽ dùng font mặc định (giới hạn ký tự Vietnamese)')


if __name__ == '__main__':
    main()
