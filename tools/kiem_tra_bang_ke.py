"""Kiem tra bang ke: voi moi ten sach + nam, in ra cac ung vien khop nhat.

Dung de ra soat tay xem tool tra co dung sach khong.

Cach chay:
    python tools/kiem_tra_bang_ke.py "H:\\tool_files\\_test_dau_vao"
    python tools/kiem_tra_bang_ke.py --ten "Giáo trình lịch sử kinh tế quốc dân" --nam 2007
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
GOC = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, GOC)
sys.path.insert(0, os.path.join(GOC, "_libs"))

import bang_ke  # noqa: E402
import naming  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="Kiem tra tra cuu bang ke sach")
    ap.add_argument("thu_muc", nargs="?", help="thu muc chua file can kiem tra")
    ap.add_argument("--ten", help="kiem tra 1 ten sach")
    ap.add_argument("--nam", type=int, help="nam xuat ban")
    ap.add_argument("--config", default=os.path.join(GOC, "config.json"))
    ap.add_argument("--top", type=int, default=3)
    args = ap.parse_args()

    cfg = json.load(open(args.config, encoding="utf-8"))
    c = cfg.get("bang_ke_sach", {})
    duong_dan = c.get("file", "")
    if not os.path.isabs(duong_dan):
        duong_dan = os.path.join(GOC, duong_dan)
    bk = bang_ke.BangKeSach(duong_dan, nguong=c.get("nguong_khop", 0.6),
                            tim_nam_lan_can=c.get("tim_nam_lan_can", 1),
                            ban_do=cfg.get("ban_bien_tap", {}))

    if args.ten:
        ds = [(args.ten, args.nam)]
    else:
        ds = []
        for root, _d, files in os.walk(args.thu_muc or "."):
            for fn in sorted(files):
                if os.path.splitext(fn)[1].lower() not in (".pdf", ".docx", ".doc"):
                    continue
                old = naming.parse_old_name(fn)
                ten = naming.normalize_base(
                    old["base"] if old.get("co_duoi") else os.path.splitext(fn)[0], cfg)
                ds.append((ten, args.nam))

    for ten, nam in ds:
        print("=" * 110)
        print("TEN SACH : %s   (nam: %s)" % (ten[:95], nam or "?"))
        for i, u in enumerate(bk.tra_nhieu(ten, nam, so_luong=args.top), 1):
            ma, gc = bk.ma_ban(u["bbt"], u["ten_sach"])
            print("  %d) diem=%.2f  BBT=%-8s -> N4=%-4s | sheet %s dong %-4s | %s"
                  % (i, u["diem"], u["bbt"] or "(trong)", ma or "??", u["sheet"], u["dong"],
                     u["ten_sach"][:70]))
            if u["tac_gia"]:
                print("       tac gia: %s" % u["tac_gia"][:80])
        kq = bk.tra(ten, nam)
        if kq:
            print("  => CHON: %s | N4=%s | dat nguong=%s chac chan=%s"
                  % (kq["ten_sach"][:60], kq["ma"], kq["dat_nguong"], kq["chac_chan"]))
    print("=" * 110)
    if bk.chua_map:
        print("Ma BBT chua map duoc:")
        for k, v in sorted(bk.chua_map.items(), key=lambda kv: -kv[1][0]):
            print("   %-12s %4d lan | vi du: %s" % (k, v[0], v[1][:60]))


if __name__ == "__main__":
    main()
