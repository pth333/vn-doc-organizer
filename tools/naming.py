"""Dat ten file theo quy dinh trong Quyet dinh 68-QD/TW.

Cau truc ten:  <TenChinh><DDmmYY><Loai><DonVi><TinhTrang><SoLanXB>.<ext>

  Nhom 1  Ten chinh file : ten sach / noi dung van ban, khong dau, noi bang "_"
  Nhom 2  6 chu so        : ngay thang nam tao file (DDmmYY)
  Nhom 3  1 chu cai       : R = file ruot, B = file bia, D = sach dien tu, K = khac
  Nhom 4  2 chu so        : don vi ban hanh (01..21)
  Nhom 5  1 chu cai       : M = xuat ban moi, T = tai ban
  Nhom 6  2 chu so        : so lan xuat ban (00 = lan dau, 01 = tai ban lan 1, ...)

Vi du: ho_chi_minh_bien_nien_tieu_su090523R04T02.PDF
"""
import os
import re
import unicodedata

# "Ten_sach160807R03M00"  ->  base, ngay, loai, don vi, tinh trang, so lan
RE_TAIL = re.compile(
    r"^(?P<base>.+?)_?(?P<date>\d{4,8})?(?P<type>[A-Za-zĐ])?"
    r"(?P<unit>\d{2})(?P<tt>[MTmt])(?P<count>\d{2})$")

LOAI_HOP_LE = {"R", "B", "D", "K"}
DON_VI_HOP_LE = {("%02d" % i) for i in range(1, 22)}
TT_HOP_LE = {"M", "T"}

# hau to rac thuong gap khi copy file
RE_RAC = re.compile(r"(\.(doc|docx|pdf|txt)\d+|\s*\(\d+\)|\s*[-_]?\s*copy)$", re.I)


def strip_accents(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def parse_old_name(filename):
    """Doc thong tin ma hoa san trong ten file cu (neu co).

    Tra ve dict: base, ngay(6 so hop le hay None), loai, don_vi, tinh_trang,
    so_lan, ngay_hop_le(bool), hop_le(bool), tho(ten goc bo phan ma hoa)
    """
    stem = os.path.splitext(filename)[0]
    # bo cac hau to rac (.doc2, (1), - Copy...)
    for _ in range(3):
        new = RE_RAC.sub("", stem)
        if new == stem:
            break
        stem = new

    m = RE_TAIL.match(stem)
    out = dict(base=stem, ngay=None, loai=None, don_vi=None, tinh_trang=None,
               so_lan=None, ngay_hop_le=False, hop_le=False, co_duoi=False)
    if not m:
        return out

    out["co_duoi"] = True
    base = m.group("base").strip(" _-")
    date, typ, unit, tt, count = (m.group("date"), m.group("type"),
                                  m.group("unit"), m.group("tt"), m.group("count"))
    # "07R03M00" -> phan "07" bi coi la base, can bo di
    if date is None:
        base = re.sub(r"[\s_\-]*\d{1,4}$", "", base).strip(" _-")
    if date and len(date) == 6:
        d, mo, y = int(date[:2]), int(date[2:4]), int(date[4:])
        if 1 <= d <= 31 and 1 <= mo <= 12:
            out["ngay"] = date
            out["ngay_hop_le"] = True
    if typ:
        typ = typ.upper()
        if typ in LOAI_HOP_LE:
            out["loai"] = typ
    if unit in DON_VI_HOP_LE:
        out["don_vi"] = unit
    if tt and tt.upper() in TT_HOP_LE:
        out["tinh_trang"] = tt.upper()
    if count:
        out["so_lan"] = count
    out["base"] = base or stem
    out["hop_le"] = all([out["ngay_hop_le"], out["loai"], out["don_vi"],
                         out["tinh_trang"], out["so_lan"] is not None])
    if not out["hop_le"]:
        # ten cu bi ma hoa loi (vi du "..._Viet_Nam_26010_R04M00"):
        # bo cac chu so dinh lien ngay sau chu cai o cuoi ten chinh
        out["base"] = re.sub(r"(?<=[A-Za-z\)])[\s_\-]?\d{1,6}$", "", out["base"]).strip(" _-")
    return out


def normalize_base(name, cfg):
    """Chuan hoa nhom 1: khong dau, khong khoang trang, noi bang '_'."""
    t = cfg.get("ten", {})
    s = name.strip()
    if t.get("bo_dau", True):
        s = strip_accents(s)
    if t.get("viet_thuong", False):
        s = s.lower()
    # bo cac hau to rac con sot lai trong ten chinh (vi du "..._QUOC_TE_doc")
    s = re.sub(r"[_\-\s]+(doc|docx|pdf|txt|copy|final|sua|in)\d*$", "", s, flags=re.I)
    # giu dau '-' trong khoang nam/thang (vi du 1950-1965, 1950_-_2006)
    s = re.sub(r"(?<=\d)[\s_]*[-\u2013][\s_]*(?=\d)", "\x00", s)
    s = re.sub(r"[\s\u00a0]+", "_", s)             # khoang trang -> _
    giu = re.escape(t.get("giu_ky_tu", "-()"))
    s = re.sub(r"[^A-Za-z0-9_%s\x00]" % giu, "_", s)   # ky tu la -> _
    s = re.sub(r"[-_]{2,}", "_", s)
    s = re.sub(r"_+", "_", s).strip("_.-")
    s = s.replace("\x00", "-")
    max_len = int(t.get("do_dai_toi_da", 170))
    if len(s) > max_len:
        s = s[:max_len].rstrip("_.-")
    return s or "khong_ten"


def guess_loai_file(text, cfg, mac_dinh):
    """Nhom 3: doan tu noi dung, khong chac thi lay mac dinh."""
    low = text[:6000].lower()
    for rule in cfg.get("doan_loai_file", []):
        for kw in rule.get("tu_khoa", []):
            if strip_accents(kw).lower() in strip_accents(low).lower():
                return rule["ma"], "tu khoa '%s' trong noi dung" % kw
    return mac_dinh, "mac dinh trong config"


def guess_don_vi(text, cfg, mac_dinh):
    """Nhom 4: doan don vi ban hanh theo tu khoa."""
    low = strip_accents(text[:8000]).lower()
    for rule in cfg.get("doan_don_vi", []):
        for kw in rule.get("tu_khoa", []):
            if strip_accents(kw).lower() in low:
                return rule["ma"], "tu khoa '%s' trong noi dung" % kw
    return mac_dinh, "mac dinh trong config"


def guess_tinh_trang(text, cfg, mac_dinh):
    """Nhom 5 + 6: M/T va so lan xuat ban."""
    low = strip_accents(text[:8000]).lower()
    rules = cfg.get("doan_tinh_trang", {})
    for kw in rules.get("tai_ban", []):
        k = strip_accents(kw).lower()
        if k in low:
            m = re.search(re.escape(k) + r"[^\n]{0,40}?(?:th[uứ]|l[ầa]n th[uứ])\s*(\d{1,2})", low)
            if m:
                n = int(m.group(1))
                return "T", "%02d" % n, "tu khoa '%s' (lan %d)" % (kw, n)
            return "T", "01", "tu khoa '%s'" % kw
    for kw in rules.get("xuat_ban_moi", []):
        if strip_accents(kw).lower() in low:
            return "M", "00", "tu khoa '%s'" % kw
    return mac_dinh, "00", "mac dinh trong config"


def build_name(base, ngay6, loai, don_vi, tinh_trang, so_lan, ext, cfg):
    ext_out = ext.upper() if cfg.get("ten", {}).get("viet_hoa_phan_mo_rong") else ext
    return "%s%s%s%s%s%s%s" % (base, ngay6, loai, don_vi, tinh_trang, so_lan, ext_out)


def unique_path(dir_path, name, da_dung=None):
    """Tra ve duong dan khong trung (them _2, _3...)."""
    da_dung = da_dung if da_dung is not None else set()
    stem, ext = os.path.splitext(name)
    cand = name
    i = 2
    while True:
        key = os.path.join(dir_path, cand).lower()
        if (key not in da_dung) and not os.path.exists(os.path.join(dir_path, cand)):
            da_dung.add(key)
            return cand
        cand = "%s_%d%s" % (stem, i, ext)
        i += 1
