"""Doc van ban (text) tu file .pdf, .docx, .doc de phan tich noi dung.

- .pdf  : pypdf (khong can cai dat gi them, thu vien nam trong _libs)
- .docx : doc word/document.xml + header/footer bang zipfile (thuan Python)
- .doc  : doc truc tiep dinh dang OLE (xem doc_reader.py), khong can MS Word
Van ban sau khi doc se duoc tu dong giai ma TCVN3 -> Unicode neu can.
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
LIBS = os.path.join(os.path.dirname(HERE), "_libs")
if os.path.isdir(LIBS) and LIBS not in sys.path:
    sys.path.insert(0, LIBS)

from doc_reader import read_doc_text  # noqa: E402
import tcvn3  # noqa: E402


class ExtractError(Exception):
    pass


# --------------------------------------------------------------------------- #
# .docx
# --------------------------------------------------------------------------- #
_TAG_RE = re.compile(r"<[^>]+>")
_P_RE = re.compile(r"</w:p>|<w:br\s*/?>|</w:tr>")


def _xml_to_text(xml):
    xml = _P_RE.sub("\n", xml)
    xml = re.sub(r"<w:tab\s*/>", "\t", xml)
    txt = _TAG_RE.sub("", xml)
    txt = (txt.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
              .replace("&quot;", '"').replace("&apos;", "'"))
    return txt


def read_docx_text(path):
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            parts = [n for n in names if n == "word/document.xml"]
            parts += sorted(n for n in names if re.match(r"word/(header|footer)\d*\.xml$", n))
            parts += sorted(n for n in names if n.startswith("word/footnotes"))
            out = []
            for n in parts:
                out.append(_xml_to_text(z.read(n).decode("utf-8", errors="replace")))
            if not out:
                raise ExtractError("file .docx khong co word/document.xml")
            return "\n".join(out)
    except zipfile.BadZipFile as e:
        raise ExtractError("file .docx hong (khong phai file zip hop le): %s" % e)


# --------------------------------------------------------------------------- #
# .pdf
# --------------------------------------------------------------------------- #
def read_pdf_text(path, max_pages=None):
    """Doc text cua PDF. max_pages=None nghia la doc TOAN BO so trang."""
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ExtractError("thieu thu vien pypdf (chay tools/download_libs.py)")
    try:
        reader = PdfReader(path)
    except Exception as e:
        raise ExtractError("khong mo duoc PDF: %s" % e)
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception:
            raise ExtractError("PDF co mat khau, khong doc duoc")
    pages = []
    n = len(reader.pages)
    ds_trang = reader.pages if not max_pages else reader.pages[:max_pages]
    for i, page in enumerate(ds_trang):
        try:
            pages.append(page.extract_text() or "")
        except Exception as e:
            pages.append("")
            if i == 0:
                raise ExtractError("khong trich duoc text tu PDF (co the la ban scan anh): %s" % e)
    text = "\n".join(pages)
    if not text.strip():
        raise ExtractError("PDF khong co lop text (ban scan anh?) - can OCR")
    return text


# --------------------------------------------------------------------------- #
# .doc
# --------------------------------------------------------------------------- #
def read_doc_text_safe(path):
    try:
        return read_doc_text(path)
    except Exception as e:
        raise ExtractError("khong doc duoc file .doc: %s" % e)


# --------------------------------------------------------------------------- #
# API chung
# --------------------------------------------------------------------------- #
def extract_text(path, max_pages=None):
    """Tra ve text cua file (da giai ma TCVN3 neu can).

    Raise ExtractError neu khong doc duoc.
    """
    ext = os.path.splitext(path)[1].lower()
    print("Dinh dang file: %s" % ext)
    if ext == ".pdf":
        text = read_pdf_text(path, max_pages=max_pages)
    elif ext == ".docx":
        text = read_docx_text(path)
    elif ext == ".doc":
        text = read_doc_text_safe(path)
    else:
        raise ExtractError("dinh dang khong duoc ho tro: %s" % ext)

    text = text.replace("\x00", "")
    text, decoded = tcvn3.decode_if_needed(text)
    return text, decoded
