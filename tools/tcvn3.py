"""Giai ma van ban TCVN3 (ABC) sang Unicode.

Nhieu file .doc cu cua cac nha xuat ban luu van ban bang font/bang ma TCVN3
(8-bit). Khi doc ra se thanh chu loi (vi du "Hµ Néi" thay vi "Hà Nội").
Module nay phat hien va giai ma lai.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tcvn3_table import PAIR, SINGLE  # noqa: E402

# Cac byte dac trung cua TCVN3 ma Unicode/Windows-1252 binh thuong hau nhu khong dung
MARKERS = set("µ¶·¸¹»¼½¾ÆÇÈÉÊËÌÎÏÐÑÒÓÔÕÖ×ØÜÝÞßåæçèéêëìíîïñòôö÷øùúûüýþ")
MARKER_STRONG = set("µ¶·¹»¼½¾ÆÇÈÉÊËÎÏÐÑÒÕÖØÜÝÞßåæçèéêëìîïñôö÷øúûüýþ")


def looks_like_tcvn3(text, min_markers=8):
    """True neu van ban co dau hieu duoc ma hoa bang TCVN3."""
    sample = text[:20000]
    if not sample:
        return False
    strong = sum(1 for ch in sample if ch in MARKER_STRONG)
    if strong >= min_markers:
        return True
    # "®" (đ), "§" (Đ), "¨" (ă), "¬" (ơ), "¦" (ư) la cac byte rat dac trung
    special = sum(1 for ch in sample if ch in "®§¨¬¦¡¤¥£¢")
    return special >= 3


def decode(text):
    """Giai ma TCVN3 -> Unicode. Cac ky tu khong thuoc bang ma giu nguyen."""
    out = []
    i = 0
    n = len(text)
    while i < n:
        if i + 1 < n:
            two = text[i:i + 2]
            rep = PAIR.get(two)
            if rep is not None:
                out.append(rep)
                i += 2
                continue
        rep = SINGLE.get(text[i])
        out.append(rep if rep is not None else text[i])
        i += 1
    return "".join(out)


def decode_if_needed(text):
    """Tra ve (van_ban, da_giai_ma)."""
    if looks_like_tcvn3(text):
        return decode(text), True
    return text, False


if __name__ == "__main__":
    tests = [
        ("Hµ Néi - 2007", "Hà Nội - 2007"),
        ("nhµ xuÊt b¶n ChÝnh trÞ quèc gia", "nhà xuất bản Chính trị quốc gia"),
        ("TS. ph¹m v¨n bÝnh", "TS. phạm văn bính"),
        ("Chñ biªn:", "Chủ biên:"),
        ("Hµ §¨ng", "Hà Đăng"),
        ("C¸c t¸c gi¶:", "Các tác giả:"),
        ("ViÖn Khoa häc ph¸p lý", "Viện Khoa học pháp lý"),
        ("®æi míi", "đổi mới"),
        ("Nhãm biªn so¹n:", "Nhóm biên soạn:"),
        ("V¨n kiÖn §¹i héi", "Văn kiện Đại hội"),
    ]
    bad = 0
    for src, want in tests:
        got = decode(src)
        flag = "OK " if got == want else "SAI"
        if got != want:
            bad += 1
        print("%s %-40s -> %s" % (flag, src, got))
    print("ket qua: %d/%d dung" % (len(tests) - bad, len(tests)))
