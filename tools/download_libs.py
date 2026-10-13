"""Tai cac thu vien Python thuan (pure-python wheels) tu PyPI vao H:\\tool_files\\_libs.

Dung khi `pip install` bi sandbox chan ghi file tam.
Chay:  python tools\\download_libs.py
"""
import os
import sys
import zipfile

import requests

LIBS = r"H:\tool_files\_libs"
WHL = os.path.join(LIBS, "_wheels")
PKGS = ["pypdf", "openpyxl", "et_xmlfile"]


def pick_wheel(files):
    for f in files:
        if f["filename"].endswith("-py3-none-any.whl"):
            return f
    for f in files:
        if f["filename"].endswith(".whl") and "none-any" in f["filename"]:
            return f
    for f in files:
        if f["filename"].endswith(".whl"):
            return f
    return None


def main():
    os.makedirs(WHL, exist_ok=True)
    for pkg in PKGS:
        meta = requests.get("https://pypi.org/pypi/%s/json" % pkg, timeout=60).json()
        ver = meta["info"]["version"]
        f = pick_wheel(meta["releases"][ver])
        if not f:
            print("!! khong tim thay wheel cho", pkg)
            continue
        dest = os.path.join(WHL, f["filename"])
        if not os.path.exists(dest):
            r = requests.get(f["url"], timeout=120)
            r.raise_for_status()
            with open(dest, "wb") as fh:
                fh.write(r.content)
        with zipfile.ZipFile(dest) as z:
            z.extractall(LIBS)
        print("OK %-12s %s" % (pkg, f["filename"]))

    sys.path.insert(0, LIBS)
    import pypdf  # noqa
    import openpyxl  # noqa
    print("IMPORT OK pypdf=%s openpyxl=%s" % (pypdf.__version__, openpyxl.__version__))


if __name__ == "__main__":
    main()
