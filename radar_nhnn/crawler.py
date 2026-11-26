#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Radar NHNN – bộ cào số liệu Ngân hàng Nhà nước
==============================================
Nguồn chính là API "headless-delivery" mà chính các trang biểu đồ của sbv.gov.vn
gọi tới (không công bố, nhưng mở cho khách vãng lai):

    /o/headless-delivery/v1.0/content-structures/<ID>/structured-contents

    137473   Tỷ giá trung tâm VND/USD (từ 2009)
    3450514  Tỷ giá tham khảo tại Cục Quản lý ngoại hối (7 ngoại tệ)
    3450260  Lãi suất bình quân + doanh số liên ngân hàng (từ 2014)
    3450482  Lãi suất NHNN quy định (tái cấp vốn, tái chiết khấu)

Hai mục không có cấu trúc dữ liệu được đọc từ HTML:
    - Kết quả đấu thầu thị trường mở (OMO)
    - Tin CPI hằng tháng trên trang /vi/cpi

Lần chạy đầu tự tải toàn bộ lịch sử; các lần sau chỉ lấy 45 ngày gần nhất rồi gộp.
Kết quả ghi vào docs/radar/data/*.json, kèm health.json ghi tình trạng từng nguồn.

Chạy:  python radar_nhnn/crawler.py            (tăng dần)
       python radar_nhnn/crawler.py --full     (tải lại toàn bộ lịch sử)
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import sys
import time
from html.parser import HTMLParser

import requests

BASE = "https://sbv.gov.vn"
API = BASE + "/o/headless-delivery/v1.0/content-structures/{sid}/structured-contents"
STRUCTURES = {
    "central": 137473,
    "reference": 3450514,
    "interbank": 3450260,
    "policy": 3450482,
}
PAGES = {
    "omo": BASE + "/vi/nghi%E1%BB%87p-v%E1%BB%A5-th%E1%BB%8B-tr%C6%B0%E1%BB%9Dng-m%E1%BB%9F",
    "cpi": BASE + "/vi/cpi",
}
LINKS = {
    "central": BASE + "/bieu-do-ty-gia-trung-tam",
    "reference": BASE + "/vi/t%E1%BB%B7-gi%C3%A1",
    "interbank": BASE + "/l%C3%A3i-su%E1%BA%A5t1",
    "policy": BASE + "/l%C3%A3i-su%E1%BA%A5t1",
    "omo": PAGES["omo"],
    "cpi": PAGES["cpi"],
}
TITLES = {
    "central": "Tỷ giá trung tâm VND/USD",
    "reference": "Tỷ giá tham khảo tại Cục QLNH",
    "interbank": "Lãi suất & doanh số liên ngân hàng",
    "policy": "Lãi suất NHNN quy định",
    "omo": "Kết quả đấu thầu thị trường mở (OMO)",
    "cpi": "CPI (tin NHNN đăng lại từ Cục Thống kê)",
}

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "docs", "radar", "data")
VN = dt.timezone(dt.timedelta(hours=7))
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept": "application/json, text/html;q=0.9, */*;q=0.8",
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.6",
}

session = requests.Session()
session.headers.update(HEADERS)


def now_vn():
    return dt.datetime.now(VN)


def log(msg):
    print(f"[{now_vn():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


class Blocked(RuntimeError):
    """Máy chủ NHNN trả 403: tường lửa đang chặn tạm vì gọi dồn dập."""


PACE_S = float(os.environ.get("RADAR_PACE_S", "20"))   # giãn cách giữa 2 lần gọi
_state = {"last": 0.0, "blocked": False, "calls": 0}


def get(url, params=None, tries=3, timeout=90):
    """GET chậm rãi, có thử lại. Gặp 403 thì dừng hẳn mọi lần gọi còn lại của lượt chạy này,
    vì gọi tiếp chỉ làm lệnh chặn kéo dài."""
    if _state["blocked"]:
        raise Blocked("bỏ qua: NHNN đang chặn tạm (403) trong lượt chạy này")
    last = None
    for i in range(tries):
        wait = _state["last"] + PACE_S - time.time()
        if wait > 0:
            time.sleep(wait)
        try:
            _state["calls"] += 1
            r = session.get(url, params=params, timeout=timeout)
            _state["last"] = time.time()
            if r.status_code == 200:
                return r
            if r.status_code in (403, 429):
                _state["blocked"] = True
                raise Blocked(f"HTTP {r.status_code}: NHNN chặn tạm vì gọi nhiều – để lượt sau")
            last = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            _state["last"] = time.time()
            last = f"{type(e).__name__}: {e}"
        log(f"  lỗi {last} – thử lại")
        time.sleep(10 * (i + 1))
    raise RuntimeError(last)


# ------------------------------------------------------------------ helpers
def vn_num(s):
    """'1.053.948,0' -> 1053948.0 ; '4,69' -> 4.69 ; '25676' -> 25676.0 ; '' -> None"""
    if s is None:
        return None
    s = str(s).strip().replace("%", "").replace("\xa0", "").replace(" ", "")
    if not s or s in {"-", "–"}:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1 or re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    if s.startswith("."):
        s = "0" + s
    try:
        return float(s)
    except ValueError:
        return None


def iso_to_vn_date(s):
    """'2026-11-25T17:00:00Z' -> '2026-11-26' (ngày theo giờ Việt Nam)."""
    if not s:
        return None
    d = dt.datetime.fromisoformat(s.replace("Z", "+00:00"))
    return d.astimezone(VN).date().isoformat()


def fields(item):
    """Trả về dict tên trường -> giá trị, và danh sách các nhóm lặp lại."""
    flat, groups = {}, {}
    for f in item.get("contentFields", []):
        nested = f.get("nestedContentFields") or []
        if nested:
            groups.setdefault(f["name"], []).append(
                {n["name"]: (n.get("contentFieldValue") or {}).get("data") for n in nested})
        else:
            flat[f["name"]] = (f.get("contentFieldValue") or {}).get("data")
    return flat, groups


def api_page(sid, flt=None, page_size=200):
    """Một trang của API, mới nhất trước. flt là bộ lọc OData của Liferay."""
    params = {"pageSize": page_size, "page": 1, "sort": "datePublished:desc",
              "fields": "id,title,datePublished,contentFields"}
    if flt:
        params["filter"] = flt
    r = get(API.format(sid=sid), params=params)
    if "json" not in r.headers.get("Content-Type", ""):
        raise RuntimeError("API trả về không phải JSON")
    j = r.json()
    return j.get("items") or [], j.get("totalCount")


def zstamp(x):
    return x.strftime("%Y-%m-%dT%H:%M:%SZ")


def fetch_structure(sid, since=None):
    """Lấy mọi bản ghi đăng từ `since` (dùng cho số liệu nhỏ hoặc lượt tăng dần)."""
    flt = f"datePublished ge {zstamp(since)}" if since else None
    items, _ = api_page(sid, flt, page_size=200)
    return items


# ------------------------------------------------------------ API parsers
def parse_central(items):
    rows = {}
    for it in items:
        f, _ = fields(it)
        d = iso_to_vn_date(f.get("NgayBatDau")) or iso_to_vn_date(it.get("datePublished"))
        v = vn_num(f.get("TyGiaSo"))
        if d and v:
            rows[d] = {"d": d, "v": v, "vb": f.get("SoVanBan") or "",
                       "pub": it.get("datePublished")}
    return rows


def parse_reference(items):
    rows = {}
    for it in items:
        f, g = fields(it)
        d = iso_to_vn_date(f.get("ngayApDung")) or iso_to_vn_date(it.get("datePublished"))
        cur = {}
        for x in g.get("tyGiaThamKhaos", []):
            code = (x.get("ngoaiTe") or "").split("-")[0].strip()
            if code:
                cur[code] = [vn_num(x.get("mua")), vn_num(x.get("ban"))]
        if d and cur:
            rows[d] = {"d": d, "fx": cur, "pub": it.get("datePublished")}
    return rows


TENORS = {"qua đêm": "ON", "1 tuần": "1W", "2 tuần": "2W", "1 tháng": "1M",
          "3 tháng": "3M", "6 tháng": "6M", "9 tháng": "9M", "12 tháng": "12M"}


def parse_interbank(items):
    rows = {}
    for it in items:
        f, g = fields(it)
        d = iso_to_vn_date(f.get("ngayApDung"))
        if not d:
            continue
        r, vol, note = {}, {}, {}
        for x in g.get("laiSuatThiTruongNganHangs", []):
            key = TENORS.get((x.get("thoihan") or "").strip().lower())
            if not key:
                continue
            r[key] = vn_num(x.get("laiSuatBQLienNganHang"))
            vol[key] = vn_num(x.get("doanhSo"))
            if (x.get("footNode") or "").strip():
                note[key] = (x.get("noiDungFootNode") or "").strip()
        if r:
            row = {"d": d, "r": r, "vol": vol, "pub": it.get("datePublished")}
            if note:
                row["note"] = note
            # nếu trùng ngày áp dụng, giữ bản đăng sau cùng
            if d not in rows or (row["pub"] or "") > (rows[d]["pub"] or ""):
                rows[d] = row
    return rows


def parse_policy(items):
    out = []
    for it in items:
        f, g = fields(it)
        for x in g.get("banglaiSuats", []):
            out.append({"loai": (x.get("loaiLaiSuat") or "").strip(),
                        "gia_tri": vn_num(x.get("giaTri")),
                        "van_ban": (x.get("vanBanQuyetDinh") or "").strip(),
                        "ap_dung": iso_to_vn_date(f.get("ngayApDung"))})
    return out


# ------------------------------------------------------------ HTML parsers
class TableGrabber(HTMLParser):
    """Lấy các bảng HTML thành list các dòng (list ô), kèm class của dòng."""

    def __init__(self, table_class):
        super().__init__(convert_charrefs=True)
        self.want = table_class
        self.depth = 0
        self.tables, self.row, self.cell = [], None, None
        self.row_cls = ""

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            if self.depth or self.want in (a.get("class") or ""):
                self.depth += 1
                if self.depth == 1:
                    self.tables.append([])
        elif self.depth == 1 and tag == "tr":
            self.row, self.row_cls = [], a.get("class") or ""
        elif self.depth == 1 and tag in ("td", "th") and self.row is not None:
            self.cell = []

    def handle_endtag(self, tag):
        if tag == "table" and self.depth:
            self.depth -= 1
        elif self.depth == 1 and tag in ("td", "th") and self.cell is not None:
            self.row.append(re.sub(r"\s+", " ", "".join(self.cell)).strip())
            self.cell = None
        elif self.depth == 1 and tag == "tr" and self.row is not None:
            self.tables[-1].append((self.row_cls, self.row))
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def parse_omo(page_html):
    m = re.search(r'class="ls01-date"[^>]*>\s*Ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})', page_html)
    if not m:
        raise RuntimeError("Không thấy ngày đấu thầu OMO")
    d = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
    tg = TableGrabber("ls01-table")
    tg.feed(page_html)
    if not tg.tables:
        raise RuntimeError("Không thấy bảng OMO")
    sessions, cur = [], None
    for cls, cells in tg.tables[0]:
        if "ls01-group" in cls:
            cur = {"loai": cells[0], "ky_han": [], "tong": None}
            sessions.append(cur)
        elif "ls01-total" in cls and cur:
            cur["tong"] = vn_num(next((c for c in cells[1:] if c), None))
        elif cur and len(cells) >= 4 and "Kỳ hạn" in cells[0]:
            days = re.search(r"(\d+)", cells[0])
            tv, tt = (cells[1].split("/") + [""])[:2]
            cur["ky_han"].append({"ngay": int(days.group(1)) if days else None,
                                  "tham_gia": vn_num(tv), "trung_thau": vn_num(tt),
                                  "kl": vn_num(cells[2]) or 0.0, "ls": vn_num(cells[3])})
    if not sessions:
        raise RuntimeError("Bảng OMO rỗng")
    for i, s in enumerate(sessions):
        s["phien"] = i + 1
    return {"d": d, "phien": sessions,
            "tong": round(sum(s["tong"] or 0 for s in sessions), 2)}


CPI_ITEM = re.compile(
    r'class="news-title">\s*(?P<title>CPI tháng (?P<m>\d{1,2})/(?P<y>\d{4})[^<]*)</a>'
    r'.*?class="top-news-detail">(?P<body>.*?)</span>.*?class="news-date">\s*\((?P<date>[\d/]+)\)',
    re.S)


def parse_cpi(page_html):
    out = {}
    for m in CPI_ITEM.finditer(page_html):
        body = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", m.group("body")))).strip()
        key = f"{m.group('y')}-{int(m.group('m')):02d}"

        mom = re.search(r"(tăng|giảm)\s*([\d,]+)%\s*so với tháng trước", body)
        yoy = re.search(r"(tăng|giảm)\s*([\d,]+)%\s*so với cùng kỳ năm trước", body)
        avg = re.search(r"[Bb]ình quân \S+ tháng[^;.]{0,40}?CPI (tăng|giảm)\s*([\d,]+)%", body)
        core = re.search(r"lạm phát cơ bản (tăng|giảm)\s*([\d,]+)%", body)
        sgn = lambda g: None if not g else (vn_num(g.group(2)) * (-1 if g.group(1) == "giảm" else 1))
        out[key] = {"thang": key, "mom": sgn(mom), "yoy": sgn(yoy), "bq": sgn(avg),
                    "co_ban_bq": sgn(core), "tieu_de": m.group("title").strip(),
                    "dang_ngay": m.group("date"), "tom_tat": body[:600]}
    if not out:
        raise RuntimeError("Không đọc được tin CPI nào")
    return out


# ------------------------------------------------------------ storage
def path(name):
    return os.path.join(OUT, name + ".json")


def load(name, default):
    try:
        with open(path(name), encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def save(name, data):
    os.makedirs(OUT, exist_ok=True)
    tmp = path(name) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path(name))


def merge_rows(old_list, new_map):
    m = {r["d"]: r for r in old_list}
    m.update(new_map)
    return [m[k] for k in sorted(m)]


# ------------------------------------------------------------ main
def run(full=False):
    health = load("health", {"sources": {}})
    started = now_vn()
    since = None if full else (dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=45))
    results = {}

    max_calls = int(os.environ.get("RADAR_MAX_CALLS", "4"))
    today = started.date().isoformat()

    def due(key):
        """Nguồn nào cần gọi ở lượt này. NHNN chỉ cho khoảng 4 lần gọi liền nhau rồi chặn,
        nên mỗi lượt chỉ gọi những nguồn có khả năng đã có số mới."""
        if full:
            return True
        src = health["sources"].get(key, {})
        try:
            ago_h = (started - dt.datetime.fromisoformat(src["kiem_tra"])).total_seconds() / 3600
        except (KeyError, TypeError, ValueError):
            return True
        if src.get("trang_thai") != "ok":
            return True
        latest = src.get("ky_moi_nhat") or ""
        if key in ("central", "reference"):      # công bố trước 9h mỗi ngày làm việc
            return latest < today or ago_h >= 6
        if key == "omo":                          # có kết quả sau phiên đấu thầu buổi sáng/trưa
            return (latest < today and started.hour >= 9) or ago_h >= 6
        if key == "interbank":                    # công bố trễ 1–2 ngày làm việc
            return ago_h >= 2
        return ago_h >= 12                        # cpi, policy

    def step(key, fn):
        src = health["sources"].setdefault(key, {})
        src.update({"ten": TITLES[key], "link": LINKS[key]})
        if not due(key):
            results[key] = src.get("trang_thai", "ok")
            return
        if _state["calls"] >= max_calls:
            log(f"…  {key}: để lượt sau (đã dùng {_state['calls']} lần gọi)")
            results[key] = src.get("trang_thai", "cho")
            src.setdefault("trang_thai", "cho")
            return
        t0 = time.time()
        src["kiem_tra"] = started.isoformat(timespec="seconds")
        try:
            info = fn()
            src.update({"trang_thai": "ok", "loi": None, "thoi_gian_s": round(time.time() - t0, 1), **info})
            log(f"OK  {key}: {info}")
        except Exception as e:  # ghi lỗi nhưng không dừng các nguồn khác
            src.update({"trang_thai": "loi", "loi": str(e)[:300], "thoi_gian_s": round(time.time() - t0, 1)})
            log(f"LỖI {key}: {e}")
        results[key] = src["trang_thai"]

    backfill = load("backfill", {})
    budget = {"pages": 0}

    def do_series(key, parser):
        def _():
            sid = STRUCTURES[key]
            cur = [] if full else load(key, [])
            bf = {} if full else backfill.get(key, {})
            # 1) phần mới: một lần gọi cho 45 ngày gần nhất
            items, _ = api_page(sid, f"datePublished ge {zstamp(since)}")
            rows = merge_rows(cur, parser(items))
            if not bf.get("oldest_pub"):
                bf["oldest_pub"] = min((it.get("datePublished") for it in items), default=None)
            backfill[key] = bf
            save(key, rows)
            save("backfill", backfill)
            return info(key, rows)
        return _

    def info(key, rows):
        bf = backfill.get(key, {})
        return {"so_ban_ghi": len(rows), "ky_moi_nhat": rows[-1]["d"] if rows else None,
                "ky_dau_tien": rows[0]["d"] if rows else None, "lich_su_day_du": bool(bf.get("done"))}

    def do_backfill():
        """Lùi dần lịch sử theo ngày đăng, vài trang mỗi lượt chạy để không bị tường lửa chặn."""
        parsers = {"central": parse_central, "reference": parse_reference, "interbank": parse_interbank}
        done_pages = 0
        for key, parser in parsers.items():
            bf = backfill.get(key, {})
            if health["sources"].get(key, {}).get("trang_thai") != "ok":
                continue
            rows = load(key, [])
            while not bf.get("done") and bf.get("oldest_pub") and budget["pages"] > 0:
                old_items, left = api_page(STRUCTURES[key], f"datePublished lt {bf['oldest_pub']}")
                budget["pages"] -= 1
                done_pages += 1
                if not old_items:
                    bf["done"] = True
                    break
                rows = merge_rows(rows, parser(old_items))
                bf["oldest_pub"] = min(it.get("datePublished") for it in old_items)
                if len(old_items) >= (left or 0):
                    bf["done"] = True
                log(f"  {key}: lùi về {bf['oldest_pub'][:10]}, còn {max(0, (left or 0) - len(old_items))} bản ghi cũ")
                backfill[key] = bf
                save(key, rows)
                save("backfill", backfill)
            health["sources"][key].update(info(key, rows))
        return done_pages

    step("central", do_series("central", parse_central))
    step("reference", do_series("reference", parse_reference))
    def do_omo():
        rec = parse_omo(get(PAGES["omo"]).text)
        cur = load("omo", [])
        rows = merge_rows(cur, {rec["d"]: rec})
        save("omo", rows)
        return {"so_ban_ghi": len(rows), "ky_moi_nhat": rows[-1]["d"], "ky_dau_tien": rows[0]["d"]}
    step("omo", do_omo)
    step("interbank", do_series("interbank", parse_interbank))



    def do_cpi():
        new = parse_cpi(get(PAGES["cpi"]).text)
        cur = {r["thang"]: r for r in load("cpi", [])}
        for k, v in new.items():  # không để bản tin thiếu số đè lên số đã có
            old = cur.get(k, {})
            cur[k] = {**old, **{kk: vv for kk, vv in v.items() if vv is not None}}
        # Tin của NHNN có tháng không ghi CPI so với cùng kỳ: bù bằng số Cục Thống kê đã đối chiếu
        try:
            with open(os.path.join(os.path.dirname(__file__), "cpi_seed.json"), encoding="utf-8") as f:
                for s in json.load(f):
                    row = cur.setdefault(s["thang"], {"thang": s["thang"]})
                    for kk in ("yoy", "co_ban_yoy"):
                        if row.get(kk) is None and s.get(kk) is not None:
                            row[kk] = s[kk]
                            row[kk + "_nguon"] = s["nguon"]
        except FileNotFoundError:
            pass
        rows = [cur[k] for k in sorted(cur)]
        save("cpi", rows)
        return {"so_ban_ghi": len(rows), "ky_moi_nhat": rows[-1]["thang"], "ky_dau_tien": rows[0]["thang"]}
    step("cpi", do_cpi)
    def do_policy():
        rows = parse_policy(fetch_structure(STRUCTURES["policy"]))
        if not rows:
            raise RuntimeError("Không có dòng lãi suất điều hành")
        save("policy", rows)
        return {"so_ban_ghi": len(rows), "ky_moi_nhat": max(r["ap_dung"] or "" for r in rows)}
    step("policy", do_policy)

    budget["pages"] = max(0, max_calls - _state["calls"])
    try:
        n = do_backfill()
        log(f"Lịch sử: đã tải thêm {n} trang")
    except Exception as e:  # bị chặn giữa chừng thì lượt sau làm tiếp từ chỗ đã dừng
        log(f"Tạm dừng tải lịch sử: {e}")
    health["so_lan_goi"] = _state["calls"]

    health["lan_chay"] = started.isoformat(timespec="seconds")
    health["xong_luc"] = now_vn().isoformat(timespec="seconds")
    health["may_chay"] = os.environ.get("RADAR_RUNNER", "github-actions" if os.environ.get("GITHUB_ACTIONS") else "may-ca-nhan")

    # Tránh commit vô ích: nếu không có gì mới và lần ghi trước chưa quá 3 giờ thì giữ nguyên health.json
    def essence(h):
        return {k: {kk: vv for kk, vv in v.items() if kk not in ("kiem_tra", "thoi_gian_s")}
                for k, v in h.get("sources", {}).items()}
    old = load("health", {})
    try:
        prev = dt.datetime.fromisoformat(old.get("lan_chay"))
        fresh = (started - prev) < dt.timedelta(hours=3)
    except (TypeError, ValueError):
        fresh = False
    if not (fresh and essence(old) == essence(health)):
        save("health", health)
    ok = sum(v == "ok" for v in results.values())
    log(f"Xong: {ok}/{len(results)} nguồn ổn")
    return ok, len(results)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true", help="tải lại toàn bộ lịch sử")
    a = ap.parse_args()
    ok, n = run(full=a.full)
    sys.exit(0 if ok else 1)
