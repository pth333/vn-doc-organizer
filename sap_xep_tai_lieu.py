#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""SAP XEP TAI LIEU THEO NAM VA DAT TEN THEO QUY DINH 68-QD/TW
================================================================

Cong cu nay se:

1. Quet toan bo file .pdf / .docx / .doc trong thu muc nguon (ke ca thu muc con).
2. Doc noi dung tung file de tim NAM (va ngay, tac gia neu co).
3. Copy file vao thu muc dich, chia theo nam:  <thu muc dich>\\<nam>\\
4. Doi ten file theo cau truc:  <TenChinh><DDmmYY><Loai><DonVi><M_T><SoLanXB>.<ext>
   (xem tools/naming.py de biet y nghia tung nhom).
5. Ghi 1 file Excel bao cao gom:
     - Da_xu_ly                    : moi file da xu ly (ten cu, ten moi, nam, tac gia...)
     - Thieu_nam_trong_noi_dung    : file khong doc duoc nam (kem ngay chinh sua file)
     - Thieu_ngay_trong_noi_dung   : file khong doc duoc ngay trong noi dung
     - Loi_doc_file                : file loi / khong doc duoc noi dung
     - Khong_ho_tro                : file khac dinh dang pdf/docx/doc
     - Tong_hop_theo_nam           : thong ke theo nam
     - Cau_hinh                    : tham so da dung khi chay

Vi du chay:
    python sap_xep_tai_lieu.py                 # dung config.json
    python sap_xep_tai_lieu.py --dry-run       # chi xem truoc, khong copy
    python sap_xep_tai_lieu.py --limit 5       # thu voi 5 file dau
"""
import argparse
import datetime as dt
import json
import os
import shutil
import sys
import time

if getattr(sys, "frozen", False):
    HERE = os.path.dirname(sys.executable)          # thư mục chứa file .exe
else:
    HERE = os.path.dirname(os.path.abspath(__file__))
    
TOOLS = os.path.join(HERE, "tools")
LIBS = os.path.join(HERE, "_libs")
for p in (TOOLS, LIBS, HERE):
    if os.path.isdir(p) and p not in sys.path:
        sys.path.insert(0, p)

from text_extract import extract_text, ExtractError  # noqa: E402
from meta_extract import analyze  # noqa: E402
import naming  # noqa: E402

MAC_DINH_CFG = "config.json"


# --------------------------------------------------------------------------- #
# Cau hinh
# --------------------------------------------------------------------------- #
def load_config(path):
    with open(path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg.setdefault("ten", {})
    cfg.setdefault("mac_dinh", {})
    return cfg


def quet_file(src, cfg):
    """Liet ke file can xu ly (duyet ca thu muc con)."""
    ho_tro = {e.lower() for e in cfg.get("dinh_dang_ho_tro", [".pdf", ".docx", ".doc"])}
    dst = os.path.abspath(cfg.get("thu_muc_dich", ""))
    bo_qua_tam = cfg.get("bo_qua_file_tam", True)
    ds, bo_qua = [], []
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if os.path.abspath(os.path.join(root, d)) != dst]
        for fn in sorted(files):
            p = os.path.join(root, fn)
            ext = os.path.splitext(fn)[1].lower()
            if bo_qua_tam and fn.startswith("~$"):
                bo_qua.append((p, "file tam cua Word (~$) - thuong la file rac"))
                continue
            if ext not in ho_tro:
                bo_qua.append((p, "dinh dang %s khong nam trong danh sach ho tro" % (ext or "(khong co)")))
                continue
            ds.append(p)
    ds.sort()
    return ds, bo_qua


# --------------------------------------------------------------------------- #
# Chon ngay (nhom 2)
# --------------------------------------------------------------------------- #
def chon_ngay(info, old, mtime, cfg, nam_thu_muc=None):
    """Tra ve (dd, mm, yy, nguon, canh_bao, khoa_nguon).

    Thu tu:
      1. Ung vien co nam trung voi nam thu muc (theo thu tu uu tien trong config).
      2. Ngay co san trong ten file cu (ngay tao file that) - giu nguyen nam,
         chi canh bao neu khac nam thu muc.
      3. Ung vien con lai, ep nam theo thu muc (co canh bao).
    """
    ung_vien = {}

    if old.get("ngay_hop_le"):
        n = old["ngay"]
        ung_vien["ten_file_cu"] = (int(n[:2]), int(n[2:4]), int(n[4:]),
                                   "6 so trong ten file cu (%s)" % n)

    d = info.get("ngay_du")
    if d:
        ung_vien["noi_dung_du_ngay"] = (d["ngay"], d["thang"], d["nam"] % 100,
                                        "noi dung: %s" % d["chuoi"][:40])

    m = info.get("ngay_thang")
    if m:
        ung_vien["noi_dung_thang_nam"] = (1, m["thang"], m["nam"] % 100,
                                          "noi dung chi co thang/nam: %s" % m["chuoi"][:40])

    if mtime:
        ung_vien["ngay_sua_file"] = (mtime.day, mtime.month, mtime.year % 100,
                                     "ngay chinh sua file tren o dia")

    def _cb(key):
        cb = []
        if key == "noi_dung_thang_nam":
            cb.append("noi dung chi co thang/nam, da lay ngay = 01")
        if key == "ngay_sua_file":
            cb.append("khong doc duoc ngay tao file, da dung ngay chinh sua file tren o dia")
        return cb

    def _cung_nam(yy):
        return (nam_thu_muc is None) or (yy + 2000 == nam_thu_muc) or (yy + 1900 == nam_thu_muc)

    thu_tu = [k for k in cfg.get("ngay_uu_tien", []) if k in ung_vien]
    phai_cung_nam = bool(cfg.get("ngay_phai_cung_nam_voi_thu_muc", True))

    # 1. uu tien ung vien cung nam voi thu muc
    if phai_cung_nam and nam_thu_muc:
        for key in thu_tu:
            dd, mm, yy, nguon = ung_vien[key]
            if _cung_nam(yy):
                return dd, mm, yy, nguon, _cb(key), key

        # 2. ngay tao file that trong ten cu -> giu nguyen, chi canh bao
        if "ten_file_cu" in ung_vien:
            dd, mm, yy, nguon = ung_vien["ten_file_cu"]
            return dd, mm, yy, nguon, [
                "ngay tao file (%02d/%02d/20%02d) khac nam xuat ban (%s) cua thu muc"
                % (dd, mm, yy, nam_thu_muc)], "ten_file_cu"

        # 3. khong con cach nao khac: ep nam theo thu muc
        if thu_tu:
            dd, mm, yy, nguon = ung_vien[thu_tu[0]]
            return dd, mm, nam_thu_muc % 100, nguon, [
                "%s khong cung nam voi thu muc (%s), da doi nam trong ten file thanh %02d"
                % (nguon, nam_thu_muc, nam_thu_muc % 100)], thu_tu[0]

    for key in thu_tu:
        dd, mm, yy, nguon = ung_vien[key]
        return dd, mm, yy, nguon, _cb(key), key
    return None, None, None, "", [], None


# --------------------------------------------------------------------------- #
# Xu ly 1 file
# --------------------------------------------------------------------------- #
def xu_ly_file(path, src, cfg):
    ten_goc = os.path.basename(path)
    ext = os.path.splitext(ten_goc)[1].lower()
    st = os.stat(path)
    tm_goc = os.path.relpath(os.path.dirname(path), src)
    rec = dict(
        ten_goc=ten_goc, duong_dan_goc=path,
        thu_muc_goc="(ngay tai thu muc nguon)" if tm_goc == "." else tm_goc,
        ext=ext,
        ngay_sua=dt.datetime.fromtimestamp(st.st_mtime),
        kich_thuoc_kb=round(st.st_size / 1024.0, 1),
        trang_thai="ok", canh_bao=[],
        ten_moi="", thu_muc_dich="",
        nam=None, nam_cach="", nam_tin_cay="", nam_bang_chung="",
        tac_gia="", tac_gia_cach="",
        ngay_ten="", nguon_ngay="", ngay_noi_dung="", ngay_ten_file_cu="",
        ngay_sua_ngan="", ten_chinh="", loai="", don_vi="", tinh_trang="",
        so_lan="", nguon_metadata="",
    )

    old = naming.parse_old_name(ten_goc)
    if old.get("ngay_hop_le"):
        n = old["ngay"]
        rec["ngay_ten_file_cu"] = "%s/%s/%s" % (n[:2], n[2:4], n[4:])

    # 1. doc noi dung
    try:
        text, _dec = extract_text(path, max_pages=int(cfg.get("so_trang_pdf_toi_da")) or None)
    except ExtractError as e:
        rec["trang_thai"] = "loi_doc"
        rec["loi"] = str(e)
        return rec
    except Exception as e:  # loi khong luong truoc
        rec["trang_thai"] = "loi_doc"
        rec["loi"] = "%s: %s" % (type(e).__name__, e)
        return rec

    info = analyze(text)

    # 2. nam -> thu muc
    rec["nam"] = info["nam"]
    rec["nam_cach"] = info["nam_cach"]
    rec["nam_tin_cay"] = info["nam_tin_cay"]
    rec["nam_bang_chung"] = info["nam_bang_chung"]
    rec["tac_gia"] = info["tac_gia"]
    rec["tac_gia_cach"] = info["tac_gia_cach"]
    d = info.get("ngay_du") or info.get("ngay_thang")
    if d:
        rec["ngay_noi_dung"] = "%02d/%02d/%04d" % (d["ngay"], d["thang"], d["nam"])

    if not info["nam"]:
        rec["trang_thai"] = "thieu_nam"
        rec["canh_bao"].append("khong doc duoc nam trong noi dung")
        ten_moi = ten_goc  # giu nguyen ten, cho nguoi dung xu ly tay
        rec["ten_moi"] = ten_moi
        rec["thu_muc_dich"] = os.path.join(cfg["thu_muc_dich"],
                                           cfg.get("thu_muc_khong_xac_dinh_nam", "_KHONG_XAC_DINH_NAM"))
        rec["ten_chinh"] = naming.normalize_base(os.path.splitext(ten_goc)[0], cfg)
        return rec

    # 3. ngay (nhom 2)
    dd, mm, yy, nguon_ngay, cb, key_ngay = chon_ngay(info, old, rec["ngay_sua"], cfg, info["nam"])
    rec["canh_bao"].extend(cb)
    if dd is None:
        dd, mm, yy = 1, 1, info["nam"] % 100
        nguon_ngay = "khong tim thay ngay nao -> dung 01/01/%02d" % yy
        rec["canh_bao"].append("khong co ngay trong noi dung va ten file, da dung 01/01")
        key_ngay = "khong_co"
    rec["ngay_ten"] = "%02d%02d%02d" % (dd, mm, yy)
    rec["nguon_ngay"] = nguon_ngay
    rec["ngay_sua_ngan"] = rec["ngay_sua"].strftime("%d/%m/%Y")

    # 4. cac nhom 3-6: uu tien thong tin co san trong ten cu
    nguon_md = []
    if old.get("loai"):
        rec["loai"], n1 = old["loai"], "ten file cu"
    else:
        rec["loai"], n1 = naming.guess_loai_file(text, cfg, cfg["mac_dinh"].get("loai_file", "R"))
        rec["canh_bao"].append("nhom 3 (loai file) phai doan: %s" % n1)
    if old.get("don_vi"):
        rec["don_vi"], n2 = old["don_vi"], "ten file cu"
    else:
        rec["don_vi"], n2 = naming.guess_don_vi(text, cfg, cfg["mac_dinh"].get("don_vi", "04"))
        rec["canh_bao"].append("nhom 4 (don vi) phai doan: %s" % n2)
    if old.get("tinh_trang") and old.get("so_lan") is not None:
        rec["tinh_trang"], rec["so_lan"] = old["tinh_trang"], old["so_lan"]
        n3 = "ten file cu"
    else:
        tt, sl, n3 = naming.guess_tinh_trang(text, cfg, cfg["mac_dinh"].get("tinh_trang", "M"))
        rec["tinh_trang"] = old.get("tinh_trang") or tt
        rec["so_lan"] = old["so_lan"] if old.get("so_lan") is not None else sl
        rec["canh_bao"].append("nhom 5-6 (M/T, so lan XB) phai doan: %s" % n3)
    nguon_md = [n1, n2, n3]
    rec["nguon_metadata"] = " | ".join(nguon_md)

    # 5. ten chinh + ten moi
    base_goc = old["base"] if old.get("co_duoi") else os.path.splitext(ten_goc)[0]
    rec["ten_chinh"] = naming.normalize_base(base_goc, cfg)
    rec["ten_moi"] = naming.build_name(rec["ten_chinh"], rec["ngay_ten"], rec["loai"],
                                       rec["don_vi"], rec["tinh_trang"], rec["so_lan"],
                                       ext, cfg)
    rec["thu_muc_dich"] = os.path.join(cfg["thu_muc_dich"], str(info["nam"]))

    # 6. canh bao chat luong
    if rec["nam_tin_cay"] == "thap":
        rec["canh_bao"].append("nam doc duoc khong chac chan (%s)" % rec["nam_cach"])
    if yy + 2000 != info["nam"] and yy + 1900 != info["nam"]:
        rec["canh_bao"].append("nam trong ten file (%02d) khac nam thu muc (%s)" % (yy, info["nam"]))
    if old.get("ngay_hop_le") and old["ngay"] != rec["ngay_ten"]:
        rec["canh_bao"].append("ngay trong ten file cu (%s) khac ngay da dung (%s)"
                               % (old["ngay"], rec["ngay_ten"]))
    if old.get("ngay_hop_le"):
        y_cu = int(old["ngay"][4:]) + (2000 if int(old["ngay"][4:]) < 80 else 1900)
        if y_cu != info["nam"]:
            rec["canh_bao"].append("nam trong ten file cu (20%02d) khac nam doc tu noi dung (%s)"
                                   % (int(old["ngay"][4:]), info["nam"]))
    return rec


# --------------------------------------------------------------------------- #
# Ghi bao cao Excel
# --------------------------------------------------------------------------- #
def ghi_excel(rows, loi, bo_qua, cfg, nguon, thoi_diem, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="2F5597")
    wrap = Alignment(vertical="top", wrap_text=True)

    def tao_sheet(ws, tieu_de, cols, data):
        ws.append(cols)
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for r in data:
            ws.append(r)
        for c in range(1, len(cols) + 1):
            ws.column_dimensions[get_column_letter(c)].width = tieu_de[c - 1]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), ws.max_row)

    # ---- Da_xu_ly
    ws = wb.active
    ws.title = "Da_xu_ly"
    cols = ["STT", "Ten file goc", "Ten file moi", "Thu muc (nam)", "Nam", "Nguon nam",
            "Do tin cay nam", "Ten chinh (nhom 1)", "Ngay (nhom 2)", "Nguon ngay",
            "Loai (N3)", "Don vi (N4)", "Tinh trang (N5)", "So lan XB (N6)",
            "Nguon metadata", "Tac gia", "Cach xac dinh tac gia", "Ngay noi dung",
            "Ngay trong ten file cu", "Ngay sua file", "Kich thuoc (KB)",
            "Thu muc goc", "Duong dan goc", "Duong dan moi", "Canh bao"]
    widths = [5, 46, 46, 12, 7, 26, 12, 40, 10, 30, 8, 8, 10, 10, 22, 26, 24, 12, 14, 12, 11, 30, 50, 50, 44]
    data = []
    for i, r in enumerate(rows, 1):
        data.append([i, r["ten_goc"], r["ten_moi"], str(r["nam"]), r["nam"],
                      r["nam_cach"], r["nam_tin_cay"], r["ten_chinh"], r["ngay_ten"],
                      r["nguon_ngay"], r["loai"], r["don_vi"], r["tinh_trang"], r["so_lan"],
                      r["nguon_metadata"], r["tac_gia"], r["tac_gia_cach"],
                      r["ngay_noi_dung"], r["ngay_ten_file_cu"],
                      r["ngay_sua"].strftime("%d/%m/%Y %H:%M"), r["kich_thuoc_kb"],
                      r["thu_muc_goc"], r["duong_dan_goc"],
                      os.path.join(r["thu_muc_dich"], r["ten_moi"]) if r["ten_moi"] else "",
                      " ; ".join(r["canh_bao"])])
    tao_sheet(ws, widths, cols, data)

    # ---- Thieu_nam_trong_noi_dung
    ws2 = wb.create_sheet("Thieu_nam_trong_noi_dung")
    cols2 = ["STT", "Ten file", "Duong dan", "Ngay chinh sua file", "Kich thuoc (KB)",
             "Nam doc duoc tu ten file cu", "Ghi chu"]
    widths2 = [5, 50, 60, 20, 12, 22, 50]
    data2 = []
    for i, r in enumerate([x for x in loi if x["trang_thai"] == "thieu_nam"], 1):
        nam_cu = ""
        o = naming.parse_old_name(r["ten_goc"])
        if o.get("ngay_hop_le"):
            nam_cu = "20%s" % o["ngay"][4:]
        data2.append([i, r["ten_goc"], r["duong_dan_goc"],
                      r["ngay_sua"].strftime("%d/%m/%Y %H:%M"), r["kich_thuoc_kb"],
                      nam_cu, " ; ".join(r["canh_bao"])])
    tao_sheet(ws2, widths2, cols2, data2)

    # ---- Thieu_ngay_trong_noi_dung
    ws3 = wb.create_sheet("Thieu_ngay_trong_noi_dung")
    cols3 = ["STT", "Ten file goc", "Duong dan", "Ngay chinh sua file",
             "Ngay da dung (nhom 2)", "Nguon ngay da dung", "Nam", "Ghi chu"]
    widths3 = [5, 50, 60, 20, 16, 34, 8, 50]
    data3 = []
    for i, r in enumerate([x for x in rows if not x["ngay_noi_dung"]], 1):
        data3.append([i, r["ten_goc"], r["duong_dan_goc"],
                      r["ngay_sua"].strftime("%d/%m/%Y %H:%M"), r["ngay_ten"],
                      r["nguon_ngay"], r["nam"], " ; ".join(r["canh_bao"])])
    tao_sheet(ws3, widths3, cols3, data3)

    # ---- Loi_doc_file
    ws4 = wb.create_sheet("Loi_doc_file")
    cols4 = ["STT", "Ten file", "Duong dan", "Ngay chinh sua file", "Kich thuoc (KB)",
             "Loai loi", "Chi tiet loi"]
    widths4 = [5, 50, 60, 20, 12, 16, 90]
    data4 = []
    for i, r in enumerate([x for x in loi if x["trang_thai"] == "loi_doc"], 1):
        data4.append([i, r["ten_goc"], r["duong_dan_goc"],
                      r["ngay_sua"].strftime("%d/%m/%Y %H:%M"), r["kich_thuoc_kb"],
                      "khong doc duoc", r.get("loi", "")])
    tao_sheet(ws4, widths4, cols4, data4)

    # ---- Khong_ho_tro
    ws5 = wb.create_sheet("Khong_ho_tro")
    cols5 = ["STT", "Ten file", "Duong dan", "Ly do"]
    widths5 = [5, 55, 70, 60]
    data5 = [[i, os.path.basename(p), p, ly] for i, (p, ly) in enumerate(bo_qua, 1)]
    tao_sheet(ws5, widths5, cols5, data5)

    # ---- Tong_hop_theo_nam
    ws6 = wb.create_sheet("Tong_hop_theo_nam")
    cols6 = ["Nam", "So file", "Dung luong (MB)", "So file do tin cay thap",
             "So file lech nam voi ten cu", "Thu muc"]
    widths6 = [8, 10, 16, 20, 24, 40]
    theo_nam = {}
    for r in rows:
        k = r["nam"]
        t = theo_nam.setdefault(k, dict(n=0, kb=0.0, low=0, lech=0))
        t["n"] += 1
        t["kb"] += r["kich_thuoc_kb"]
        if r["nam_tin_cay"] == "thap":
            t["low"] += 1
        if any("khac nam xuat ban" in c or "khong cung nam voi thu muc" in c
               for c in r["canh_bao"]):
            t["lech"] += 1
    data6 = []
    for k in sorted(theo_nam, key=lambda x: (x is None, x)):
        t = theo_nam[k]
        data6.append([k if k else "(khong xac dinh)", t["n"], round(t["kb"] / 1024.0, 2),
                      t["low"], t["lech"],
                      os.path.join(cfg["thu_muc_dich"], str(k)) if k else
                      os.path.join(cfg["thu_muc_dich"], cfg.get("thu_muc_khong_xac_dinh_nam", ""))])
    tao_sheet(ws6, widths6, cols6, data6)

    # ---- Cau_hinh
    ws7 = wb.create_sheet("Cau_hinh")
    ws7.append(["Tham so", "Gia tri"])
    ws7["A1"].font = head_font
    ws7["A1"].fill = head_fill
    ws7["B1"].font = head_font
    ws7["B1"].fill = head_fill
    ws7.column_dimensions["A"].width = 34
    ws7.column_dimensions["B"].width = 100
    for k in ("thu_muc_nguon", "thu_muc_dich", "dinh_dang_ho_tro", "ngay_uu_tien",
              "copy_giu_file_goc", "thu_muc_khong_xac_dinh_nam"):
        ws7.append([k, json.dumps(cfg.get(k), ensure_ascii=False)])
    ws7.append(["thoi_diem_chay", thoi_diem.strftime("%d/%m/%Y %H:%M:%S")])
    ws7.append(["tong so file xu ly", len(rows)])
    ws7.append(["tong so file loi", len([x for x in loi if x["trang_thai"] == "loi_doc"])])
    ws7.append(["tong so file thieu nam", len([x for x in loi if x["trang_thai"] == "thieu_nam"])])
    ws7.append(["tong so file bo qua", len(bo_qua)])
    for row in ws7.iter_rows(min_row=2):
        row[1].alignment = wrap

    wb.save(out_path)
    return out_path


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def kiem_tra_thu_vien():
    """Bao loi ro rang neu thieu thu vien can thiet."""
    thieu = []
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        thieu.append("openpyxl (ghi bao cao Excel)")
    try:
        import pypdf  # noqa: F401
    except ImportError:
        thieu.append("pypdf (doc file PDF)")
    if thieu:
        print("!! Thieu thu vien: %s" % ", ".join(thieu))
        print("   Cach xu ly:  python tools\\download_libs.py")
        print("   hoac:        pip install pypdf openpyxl")
        return False
    return True

def _hoi(cau_hoi):
    return input(cau_hoi).strip().strip('"').strip()

def main():
    ap = argparse.ArgumentParser(description="Sap xep tai lieu theo nam va dat ten theo QD 68-QD/TW")
    ap.add_argument("--config", default=os.path.join(HERE, MAC_DINH_CFG), help="file cau hinh JSON")
    ap.add_argument("--input", help="thu muc nguon (ghi de config)")
    ap.add_argument("--output", help="thu muc dich (ghi de config)")
    ap.add_argument("--dry-run", action="store_true", help="chi xem truoc, khong copy file")
    ap.add_argument("--move", action="store_true", help="di chuyen file thay vi copy")
    ap.add_argument("--limit", type=int, default=0, help="chi xu ly N file dau (de thu)")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="in ra tung file: ten cu -> ten moi, nam, tac gia")
    args = ap.parse_args()

    if not kiem_tra_thu_vien():
        return 3

    cfg = load_config(args.config)
    
    if not cfg.get("thu_muc_nguon"):
        cfg["thu_muc_nguon"] = _hoi("Kéo thả thư mục nguồn (thư mục chứa file cần sắp xếp) vào đây rồi nhấn Enter: ")
    if not cfg.get("thu_muc_dich"):
        cfg["thu_muc_dich"] = _hoi("Kéo thả thư mục đích (thư mục để chứa file đã sắp xếp) vào đây rồi nhấn Enter: ")

    if args.input:
        cfg["thu_muc_nguon"] = args.input
    if args.output:
        cfg["thu_muc_dich"] = args.output
    if args.move:
        cfg["copy_giu_file_goc"] = False
    if args.limit:
        cfg["gioi_han_so_file"] = args.limit

    src = os.path.abspath(cfg["thu_muc_nguon"])
    dst = os.path.abspath(cfg["thu_muc_dich"])
    if not os.path.isdir(src):
        print("!! Khong ton tai thu muc nguon: %s" % src)
        return 2
    if args.dry_run:
        print(">>> CHE DO XEM TRUOC (dry-run): khong copy file nao.\n")

    ds_file, bo_qua = quet_file(src, cfg)
    if cfg.get("gioi_han_so_file"):
        ds_file = ds_file[:int(cfg["gioi_han_so_file"])]

    print("Thu muc nguon : %s" % src)
    print("Thu muc dich  : %s" % dst)
    print("So file can xu ly: %d   |   bo qua: %d\n" % (len(ds_file), len(bo_qua)))

    t0 = time.time()
    rows, loi = [], []
    da_dung = set()
    for i, p in enumerate(ds_file, 1):
        rel = os.path.relpath(p, src)
        try:
            rec = xu_ly_file(p, src, cfg)
        except Exception as e:
            rec = dict(ten_goc=os.path.basename(p), duong_dan_goc=p, thu_muc_goc="",
                       ext=os.path.splitext(p)[1], kich_thuoc_kb=0,
                       ngay_sua=dt.datetime.fromtimestamp(os.path.getmtime(p)),
                       trang_thai="loi_doc", loi="%s: %s" % (type(e).__name__, e),
                       canh_bao=[], ten_moi="", thu_muc_dich="", nam=None, nam_cach="",
                       nam_tin_cay="", nam_bang_chung="", tac_gia="", tac_gia_cach="",
                       ngay_ten="", nguon_ngay="", ngay_noi_dung="", ngay_ten_file_cu="",
                       ngay_sua_ngan="", ten_chinh="", loai="", don_vi="", tinh_trang="",
                       so_lan="", nguon_metadata="")
        if rec["trang_thai"] != "ok":
            loi.append(rec)
            print("[%3d/%3d] %-8s %s" % (i, len(ds_file), rec["trang_thai"], rel[:70]))
            if rec["trang_thai"] == "thieu_nam":
                # van copy sang thu muc _KHONG_XAC_DINH_NAM voi ten goc
                if not args.dry_run:
                    os.makedirs(rec["thu_muc_dich"], exist_ok=True)
                    ten = naming.unique_path(rec["thu_muc_dich"], rec["ten_moi"], da_dung)
                    rec["ten_moi"] = ten
                    shutil.copy2(p, os.path.join(rec["thu_muc_dich"], ten))
            continue

        # dat ten duy nhat va copy
        os.makedirs(rec["thu_muc_dich"], exist_ok=True) if not args.dry_run else None
        ten = naming.unique_path(rec["thu_muc_dich"], rec["ten_moi"], da_dung)
        if ten != rec["ten_moi"]:
            rec["canh_bao"].append("ten bi trung, da them so thu tu: %s" % ten)
            rec["ten_moi"] = ten
        if not args.dry_run:
            dich = os.path.join(rec["thu_muc_dich"], rec["ten_moi"])
            if os.path.exists(dich):
                if cfg.get("ghi_de_file_dich"):
                    os.remove(dich)
                else:
                    rec["canh_bao"].append("file dich da ton tai, bo qua copy")
                    rows.append(rec)
                    continue
            if cfg.get("copy_giu_file_goc", True):
                shutil.copy2(p, dich)
            else:
                shutil.move(p, dich)
        rows.append(rec)

        if i % 10 == 0 or i == len(ds_file):
            print("  ... da xu ly %d/%d file" % (i, len(ds_file)))

        if args.verbose:
            print("[%3d] %-58s -> %s\\%s  | nam=%s (%s) | tg=%s" % (
                i, rec["ten_goc"][:58], rec["nam"], rec["ten_moi"][:60],
                rec["nam"], rec["nam_tin_cay"], rec["tac_gia"][:28]))

    # ---- bao cao
    thoi_diem = dt.datetime.now()
    if not args.dry_run:
        os.makedirs(dst, exist_ok=True)
        out = os.path.join(dst, "Bao_cao_sap_xep_%s.xlsx" % thoi_diem.strftime("%Y%m%d_%H%M%S"))
        ghi_excel(rows, loi, bo_qua, cfg, src, thoi_diem, out)
    else:
        out = "(dry-run: khong ghi bao cao)"

    print("\n================ KET QUA ================")
    print("Xu ly thanh cong      : %d file" % len(rows))
    print("Khong doc duoc nam    : %d file" % len([x for x in loi if x["trang_thai"] == "thieu_nam"]))
    print("Loi khong doc duoc    : %d file" % len([x for x in loi if x["trang_thai"] == "loi_doc"]))
    print("Bo qua (khac dinh dang): %d file" % len(bo_qua))
    theo_nam = {}
    for r in rows:
        theo_nam[r["nam"]] = theo_nam.get(r["nam"], 0) + 1
    for k in sorted(theo_nam, key=lambda x: (x is None, x)):
        print("   - Nam %-6s: %d file" % (k, theo_nam[k]))
    print("Bao cao Excel         : %s" % out)
    print("Thoi gian chay        : %.1f giay" % (time.time() - t0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
