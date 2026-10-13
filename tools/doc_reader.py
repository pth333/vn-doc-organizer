"""Doc text tu file .doc (Word 97-2003, dinh dang OLE/CFB) khong can Microsoft Word.

Thuat toan: doc FIB -> CLX (piece table) tu stream 0Table/1Table -> ghep cac piece text.
Tham chieu: MS-DOC [MS-DOC] 2.4.1 (FIB), 2.8.1 (Piece Table).
"""
import struct

import olefile


def _decode_guess(raw):
    """Doan encoding cho truong hop file .doc khong co piece table (Word 6/95)."""
    if raw.count(b"\x00") > len(raw) // 4:
        try:
            return raw.decode("utf-16-le", errors="replace")
        except Exception:
            pass
    return raw.decode("cp1252", errors="replace")


def _parse_clx(clx):
    """Tra ve bytes cua Pcdt (PlcPcd) hoac None."""
    i = 0
    n = len(clx)
    while i < n:
        t = clx[i]
        if t == 1:  # Prc
            if i + 3 > n:
                return None
            cb = struct.unpack_from("<h", clx, i + 1)[0]
            i += 3 + cb
        elif t == 2:  # Pcdt
            if i + 5 > n:
                return None
            lcb = struct.unpack_from("<I", clx, i + 1)[0]
            return clx[i + 5:i + 5 + lcb]
        else:
            return None
    return None


def read_doc_text(path):
    """Doc toan bo text cua file .doc. Raise exception neu khong doc duoc."""
    ole = olefile.OleFileIO(path)
    try:
        if not ole.exists("WordDocument"):
            raise ValueError("khong co stream WordDocument (khong phai file .doc cua Word?)")
        wd = ole.openstream("WordDocument").read()
        if len(wd) < 0x200:
            raise ValueError("stream WordDocument qua ngan")

        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        which_tbl = 1 if (flags & 0x0200) else 0
        candidates = ["1Table", "0Table"] if which_tbl else ["0Table", "1Table"]
        tbl_name = next((c for c in candidates if ole.exists(c)), None)

        fc_clx = lcb_clx = 0
        if tbl_name:
            tbl = ole.openstream(tbl_name).read()
            n_fib = struct.unpack_from("<H", wd, 0x02)[0]
            if n_fib >= 105:
                fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x01A2)

        pieces = []
        if tbl_name and lcb_clx > 0 and fc_clx + lcb_clx <= len(tbl):
            pcdt = _parse_clx(tbl[fc_clx:fc_clx + lcb_clx])
            if pcdt and len(pcdt) >= 4:
                n = (len(pcdt) - 4) // 12
                if n > 0:
                    cps = struct.unpack_from("<%dI" % (n + 1), pcdt, 0)
                    for k in range(n):
                        off = 4 * (n + 1) + k * 8
                        fc = struct.unpack_from("<I", pcdt, off + 2)[0]
                        cch = cps[k + 1] - cps[k]
                        if cch <= 0:
                            continue
                        if fc & 0x40000000:  # nen 8-bit (cp1252)
                            fc2 = (fc & ~0x40000000) // 2
                            raw = wd[fc2:fc2 + cch]
                            pieces.append(raw.decode("cp1252", errors="replace"))
                        else:  # UTF-16LE
                            raw = wd[fc:fc + cch * 2]
                            pieces.append(raw.decode("utf-16-le", errors="replace"))

        if not pieces:  # du phong: lay tu fcMin..fcMac
            fc_min, fc_mac = struct.unpack_from("<II", wd, 0x18)
            if fc_mac <= fc_min or fc_mac > len(wd):
                raise ValueError("khong xac dinh duoc vung text")
            pieces.append(_decode_guess(wd[fc_min:fc_mac]))

        text = "".join(pieces)
        # don sach ky tu dieu khien cua Word
        text = (text.replace("\r", "\n").replace("\x07", "\n")
                    .replace("\x0b", "\n").replace("\x0c", "\n")
                    .replace("\x01", "").replace("\x02", "")
                    .replace("\x13", "").replace("\x14", "").replace("\x15", "")
                    .replace("\x1e", "-").replace("\xa0", " "))
        return text
    finally:
        ole.close()


if __name__ == "__main__":
    import os
    import sys

    src = sys.argv[1] if len(sys.argv) > 1 else r"H:\tool_files\2007 file sách hoàn thành"
    files = sorted(f for f in os.listdir(src) if f.lower().endswith(".doc"))
    ok = bad = 0
    for f in files:
        try:
            t = read_doc_text(os.path.join(src, f))
            has07 = "2007" in t
            print("OK  %-6d %-5s %s" % (len(t), has07, f[:70]))
            ok += 1
        except Exception as e:
            print("ERR %-5s %s -> %s" % (type(e).__name__, f[:70], e))
            bad += 1
    print("TONG: ok=%d loi=%d" % (ok, bad))
