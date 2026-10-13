"""Kiem tra thu ket qua trich xuat nam / ngay / tac gia tren toan bo file."""
import os
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))

from text_extract import extract_text  # noqa: E402
from meta_extract import analyze  # noqa: E402

SRC = sys.argv[1] if len(sys.argv) > 1 else r"H:\tool_files\2007 file sách hoàn thành"


def main():
    t0 = time.time()
    files = []
    for root, _dirs, names in os.walk(SRC):
        for n in names:
            if n.startswith("~$"):
                continue
            files.append(os.path.join(root, n))
    files.sort()
    print("Tong so file:", len(files))
    for p in files:
        rel = os.path.relpath(p, SRC)
        try:
            text, dec = extract_text(p)
        except Exception as e:
            print("LOI  %-60s | %s" % (rel[:60], e))
            continue
        try:
            info = analyze(text)
        except Exception:
            print("LOI-PT %-58s" % rel[:58])
            traceback.print_exc()
            continue
        d = info["ngay_du"] or info["ngay_thang"]
        dstr = ("%02d/%02d/%04d (%s)" % (d["ngay"], d["thang"], d["nam"], d["kieu"])) if d else "-"
        print("%-52s | nam=%-6s %-26s | ngay=%-36s | tg=%-38s | %s" % (
            rel[:52], info["nam"], info["nam_cach"][:26], dstr[:36],
            (info["tac_gia"] or "")[:38], info["tac_gia_cach"][:28]))
    print("Thoi gian: %.1fs" % (time.time() - t0))


if __name__ == "__main__":
    main()
