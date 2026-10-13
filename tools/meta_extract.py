"""Trich xuat thong tin tu noi dung van ban: nam, ngay, tac gia, nha xuat ban."""
import re

# --------------------------------------------------------------------------- #
# Nam
# --------------------------------------------------------------------------- #
YEAR_RE = re.compile(r"\b(1[89]\d\d|20[0-4]\d)\b")

# "Ha Noi - 2007" tren trang bia sach
RE_TITLEPAGE_YEAR = re.compile(
    r"H[àa]\s*N[ộo]i\s*[-–—:]\s*((?:19|20)\d{2})", re.I)

# "xuat ban nam 2007", "in nam 2007", "nam 2007 Nha xuat ban ... xuat ban"
RE_PUB_YEAR = re.compile(
    r"(?:xu[ấa]t\s*b[ảa]n|t[áa]i\s*b[ảa]n|in)\s*(?:l[ầa]n\s*th[ứu]\s*\w+\s*)?(?:v[àa]o\s*)?n[ăa]m\s*((?:19|20)\d{2})", re.I)
RE_YEAR_PUB = re.compile(
    r"n[ăa]m\s*((?:19|20)\d{2})[^.\n]{0,40}?(?:Nh[àa]\s*xu[ấa]t\s*b[ảa]n|NXB)", re.I)
# "... (ngay 06/7/1967 - 06/7/2007), Nha xuat ban ... xuat ban lan dau ..."
RE_YEAR_BEFORE_PUB = re.compile(
    r"((?:19|20)\d{2})[^\n]{0,140}?(?:Nh[àa]\s*xu[ấa]t\s*b[ảa]n|NXB|xu[ấa]t\s*b[ảa]n)", re.I)
RE_PUBLISHER = re.compile(r"(Nh[àa]\s*xu[ấa]t\s*b[ảa]n[^\n]{0,60})", re.I)

# ngay/thang/nam
RE_D_NTN = re.compile(
    r"ng[àa]y\s+(\d{1,2})\s*(?:th[áa]ng\s+(\d{1,2})\s*)?(?:n[ăa]m\s+)?(\d{4})", re.I)
RE_D_SLASH = re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})\b")
RE_TN = re.compile(r"th[áa]ng\s+(\d{1,2})\s*n[ăa]m\s+(\d{4})", re.I)
RE_TN_ORDER = re.compile(r"n[ăa]m\s+(\d{4})\s*[,\-]?\s*th[áa]ng\s+(\d{1,2})", re.I)

MIN_YEAR, MAX_YEAR = 1900, 2035


def _valid_date(d, m, y):
    return (1 <= d <= 31) and (1 <= m <= 12) and (MIN_YEAR <= y <= MAX_YEAR)


def _norm_ws(text):
    return re.sub(r"[ \t\xa0]+", " ", text)


def extract_year(text, head_chars=8000):
    """Tra ve (nam, cach_xac_dinh, do_tin_cay, chi_tiet).

    do_tin_cay: 'cao' | 'vua' | 'thap'
    """
    head = text[:head_chars]
    head_short = text[:3000]

    m = RE_TITLEPAGE_YEAR.search(head_short) or RE_TITLEPAGE_YEAR.search(head)
    if m:
        return int(m.group(1)), "trang bia (Ha Noi - nam)", "cao", m.group(0).strip()

    m = RE_YEAR_PUB.search(head) or RE_PUB_YEAR.search(head)
    if m:
        return int(m.group(1)), "loi nha xuat ban / noi dung xuat ban", "cao", m.group(0).strip()

    m = RE_YEAR_BEFORE_PUB.search(head)
    if m:
        return int(m.group(1)), "nam gan voi cum 'Nha xuat ban/xuat ban'", "vua", m.group(0).strip()[:120]

    # "thang M nam YYYY" trong phan dau sach (thuong la ngay cua Loi nha xuat ban)
    m = RE_TN.search(head)
    if m:
        return int(m.group(2)), "ngay thang trong phan dau sach", "vua", m.group(0).strip()

    # nam xuat hien nhieu nhat trong phan dau sach
    counts = {}
    for mm in YEAR_RE.finditer(head):
        y = int(mm.group(1))
        counts[y] = counts.get(y, 0) + 1
    if counts:
        y, c = max(counts.items(), key=lambda kv: (kv[1], -kv[0]))
        if c >= 3:
            return y, "nam xuat hien nhieu nhat (phan dau sach)", "thap", "%d lan" % c

    counts = {}
    for mm in YEAR_RE.finditer(text):
        y = int(mm.group(1))
        counts[y] = counts.get(y, 0) + 1
    if counts:
        y, c = max(counts.items(), key=lambda kv: (kv[1], -kv[0]))
        return y, "nam xuat hien nhieu nhat (toan van)", "thap", "%d lan" % c

    return None, "khong tim thay nam trong noi dung", "khong", ""


def find_dates(text, head_chars=12000, prefer_year=None):
    """Tim cac ngay co trong phan dau van ban.

    Tra ve danh sach dict: {ngay, thang, nam, kieu, tin_cay, vi_tri, chuoi}
    """
    head = _norm_ws(text[:head_chars])
    out = []

    for m in RE_D_NTN.finditer(head):
        d, mo, y = int(m.group(1)), m.group(2), int(m.group(3))
        if mo is None:
            continue
        mo = int(mo)
        if _valid_date(d, mo, y):
            out.append(dict(ngay=d, thang=mo, nam=y, kieu="ngay thang nam", day_known=True,
                            tin_cay="vua", vi_tri=m.start(), chuoi=m.group(0).strip()))

    for m in RE_D_SLASH.finditer(head):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if _valid_date(d, mo, y):
            out.append(dict(ngay=d, thang=mo, nam=y, kieu="dd/mm/yyyy", day_known=True,
                            tin_cay="vua", vi_tri=m.start(), chuoi=m.group(0)))

    for m in RE_TN.finditer(head):
        mo, y = int(m.group(1)), int(m.group(2))
        if _valid_date(1, mo, y):
            out.append(dict(ngay=1, thang=mo, nam=y, kieu="thang nam (thieu ngay)",
                            day_known=False, tin_cay="thap", vi_tri=m.start(),
                            chuoi=m.group(0).strip()))

    for m in RE_TN_ORDER.finditer(head):
        y, mo = int(m.group(1)), int(m.group(2))
        if _valid_date(1, mo, y):
            out.append(dict(ngay=1, thang=mo, nam=y, kieu="nam thang (thieu ngay)",
                            day_known=False, tin_cay="thap", vi_tri=m.start(),
                            chuoi=m.group(0).strip()))

    # uu tien: dung nam tai lieu -> co du ngay -> xuat hien som nhat
    out.sort(key=lambda c: (0 if (prefer_year and c["nam"] == prefer_year) else 1,
                            0 if c["day_known"] else 1,
                            c["vi_tri"]))
    note = ""
    if prefer_year and out and out[0]["nam"] == prefer_year:
        note = "co ngay trung voi nam tai lieu"
    return out, note


def find_date_in_whole_text(text, prefer_year=None):
    cands, note = find_dates(text, head_chars=len(text), prefer_year=prefer_year)
    return cands, note


# --------------------------------------------------------------------------- #
# Tac gia
# --------------------------------------------------------------------------- #
AUTHOR_LABELS = [
    "chủ biên", "đồng chủ biên", "tác giả", "các tác giả", "nhóm tác giả",
    "nhóm biên soạn", "biên soạn", "chủ nhiệm", "người dịch", "dịch giả",
    "thể hiện", "sưu tầm", "biên tập",
]
RE_LABEL = re.compile(
    r"^\W*(?:ban |nhóm |tổ |đồng |các |chủ )?(%s)\s*[:\-–]?\s*(.*)$"
    % "|".join(re.escape(x) for x in AUTHOR_LABELS), re.I)
RE_LABEL_TAIL = re.compile(
    r"\(?\s*(%s)\s*\)?\s*$" % "|".join(re.escape(x) for x in AUTHOR_LABELS), re.I)

TITLE_WORDS = [
    "TS", "PGS", "GS", "ThS", "CN", "KS", "BS", "NSND", "NSƯT", "Luật gia",
    "Thạc sĩ", "Tiến sĩ", "Phó giáo sư", "Giáo sư", "Nhà giáo", "Nhà báo",
    "Bác sĩ", "Kỹ sư", "Cử nhân",
]
# cac cum tu mo dau thuong gap o ten sach / tieu de, khong the la ten tac gia
STOP_LINE = re.compile(
    r"^(lời|mục|phần|chương|bài|tập|nhà xuất bản|nxb|hà nội|in |trang|bìa|"
    r"mở đầu|kết luận|phụ lục|tài liệu|danh mục|bảng|hình|sách|công trình|"
    r"giáo trình|luật |pháp lệnh|quy định|quy trình|quy chế|hỏi\s*[-–]\s*đáp|"
    r"văn bản|điều lệ|cộng hòa|socialist|chú dẫn|những |lịch sử|"
    r"cẩm nang|nhân tố|phương pháp|về |bác hồ|hồ chí minh|tôn giáo|"
    r"đảng cộng sản|chủ nghĩa|học viện|viện |trường |ủy ban|bộ |sở |"
    r"công tác|tuyên giáo|thông tin|khoa học|kinh tế|chính trị|an ninh|"
    r"đổi mới|đất nước|con người|gia đình|thanh niên|phụ nữ|nông nghiệp)", re.I)
NAME_OK = re.compile(r"^[A-Za-zÀ-ỹĐđ][A-Za-zÀ-ỹĐđ\.\-\s']{1,80}$")
# tu/cum tu khong bao gio dung trong ten nguoi
NON_NAME_WORDS = set("""
và với trong của cho tại theo về trên dưới từ đến là có không do bởi để khi nếu như
các một những mọi mỗi này kia đó số lần thứ nhất hai ba bốn năm
người nước đảng nhà nghiệp đoàn hội ban ủy uỷ bộ sở tỉnh huyện xã phường thành phố quận
trung ương chính phủ quốc văn bản hướng dẫn thi hành dân vận giáo trình sách báo tạp chí
khoa học kinh tế doanh quản lý tổ chức hoạt động công tác nghiên cứu phát triển xây dựng
bảo vệ đổi mới chỉ đạo lãnh vấn đề nghiệm thực tiễn luận tư tưởng chủ nghĩa xã lịch sử
tôn giáo tín ngưỡng pháp luật điều lệ nghị định thông quyết chương phần mục lời nói đầu
giới thiệu kết phụ lục tài liệu tham khảo bài giảng cẩm nang tập tuyển toàn
đại tướng danh nhân anh hùng liệt sĩ thương binh gia đình thanh niên phụ nữ nông thôn
thị trường ngân hàng tiền tệ đầu thương mại dịch vụ sản xuất phòng an ninh ngoại giao
y tế dục đào tạo thể thao du lịch môi năng lượng giao thông bưu viễn nghệ tin
chú dẫn biên soạn sưu tầm dịch giả thể hiện chủ nhiệm
bác đồng trường chinh kiểm soát hỏi đáp phố thái vào đã
hồ chí minh lệnh religion religious policies policy regard affairs
and for the of in on to with by
socialist republic vietnam government committee
""".split())
CONNECTIVE = set("và với trong của cho tại theo về trên dưới từ đến do bởi để khi nếu như "
                 "các một những mọi mỗi này kia đó".split())


def _looks_like_person(line):
    """Kiem tra nghiem ngat: dong nay co phai ten nguoi khong."""
    s = line.strip().strip("-–—•*:,")
    if not s or len(s) > 60 or any(ch.isdigit() for ch in s):
        return False
    if ":" in s:
        return False
    if any(ch.isupper() for ch in s) is False and s.islower() is False:
        pass
    cleaned = _clean_person(s)
    if not cleaned or not NAME_OK.match(cleaned):
        return False
    words = cleaned.split()
    if not (2 <= len(words) <= 4):
        return False
    low = [w.lower() for w in words]
    if low[0] in CONNECTIVE:
        return False
    hits = sum(1 for w in low if w in NON_NAME_WORDS)
    return hits < max(1, len(low)) and hits / len(low) < 0.6
# "do tien si Nguyen Van A chu bien", "cua PGS. Tran Van B"
RE_AUTHOR_IN_TEXT = re.compile(
    r"(?:do|của|cuả)\s+((?:TS|PGS|GS|ThS|CN|KS|Luật gia|tiến sĩ|thạc sĩ|phó giáo sư|"
    r"giáo sư|nhà báo|nhà văn|nhà thơ|nhà nghiên cứu)\.?\s+"
    r"[A-ZÀ-Ỹ][^,.;:\n()]{2,50}?)\s+(?:chủ biên|biên soạn|dịch|viết|sưu tầm)", re.I)


def _clean_person(line):
    s = line.strip(" \t-–—•*:,")
    # bo hoc ham/hoc vi o dau
    for t in sorted(TITLE_WORDS, key=len, reverse=True):
        s = re.sub(r"^%s\.?\s*[,\-]?\s*" % re.escape(t), "", s, flags=re.I)
    # bo nhan vai tro dinh o cuoi ("... chu bien", "... chu nhiem")
    for t in sorted(AUTHOR_LABELS + ["đề tài", "nội dung", "bản thảo"], key=len, reverse=True):
        s = re.sub(r"[\s,;-]+%s$" % re.escape(t), "", s, flags=re.I)
    s = re.sub(r"\(.*?\)", "", s).strip(" ,;.-")
    s = re.sub(r"\s+", " ", s)
    # bo hoc ham/hoc vi con sot lai sau khi cat nhan
    for t in sorted(TITLE_WORDS, key=len, reverse=True):
        s = re.sub(r"^%s\.?\s*[,\-]?\s*" % re.escape(t), "", s, flags=re.I)
    return s.strip(" ,;.-")


def _looks_like_name(line):
    s = line.strip()
    if not s or len(s) > 90 or any(ch.isdigit() for ch in s):
        return False
    if STOP_LINE.match(s) or ":" in s:
        return False
    cleaned = _clean_person(s)
    if not cleaned or not NAME_OK.match(cleaned):
        return False
    words = cleaned.split()
    return 2 <= len(words) <= 8


def extract_author(text, head_chars=4000):
    """Tra ve (tac_gia, cach_xac_dinh). tac_gia = '' neu khong tim thay."""
    head = text[:head_chars]
    lines = [l.strip() for l in head.split("\n")]
    lines = [l for l in lines if l]

    # 1. Dong co nhan "Chu bien:", "Tac gia:" ...
    for i, line in enumerate(lines):
        m = RE_LABEL.match(line)
        if not m:
            continue
        label = m.group(1).lower()
        rest = m.group(2).strip()
        names = []
        if rest and _looks_like_name(rest):
            names.append(_clean_person(rest))
        j = i + 1
        while j < len(lines) and len(names) < 6:
            nxt = lines[j]
            if RE_LABEL.match(nxt) or not _looks_like_name(nxt):
                break
            names.append(_clean_person(nxt))
            j += 1
        if names:
            return "; ".join(dict.fromkeys(names)), "nhan '%s' trong phan dau sach" % label

    # 2. Ten nam ngay tren dong "(Chu bien)" / "(Tac gia)"
    for i, line in enumerate(lines):
        if RE_LABEL_TAIL.match(line) and len(line) <= 40 and i > 0:
            prev = lines[i - 1]
            if _looks_like_name(prev):
                return _clean_person(prev), "ten nam tren dong '(chu bien/tac gia)'"

    # 3. "do tien si Nguyen Van A chu bien" trong loi nha xuat ban
    m = RE_AUTHOR_IN_TEXT.search(head)
    if m:
        s = _clean_person(m.group(1))
        if _looks_like_name(s):
            return s, "cau 'do ... chu bien/bien soan' trong phan dau sach"

    # 4. Dong dau tien cua tai lieu neu giong ten nguoi (do tin cay thap)
    for line in lines[:6]:
        if _looks_like_person(line):
            return _clean_person(line), "dong dau tien cua tai lieu (do tin cay thap)"

    return "", "khong tim thay tac gia"


# --------------------------------------------------------------------------- #
# Nha xuat ban (tham khao)
# --------------------------------------------------------------------------- #
def extract_publisher(text, head_chars=4000):
    m = RE_PUBLISHER.search(text[:head_chars])
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" ,.-")
    return ""


def analyze(text):
    """Phan tich tong hop: tra ve dict thong tin trich xuat tu noi dung."""
    year, year_how, year_conf, year_ev = extract_year(text)
    dates, dates_note = find_dates(text, prefer_year=year)
    if not dates:
        dates, dates_note = find_date_in_whole_text(text, prefer_year=year)
        if dates:
            dates_note = (dates_note + " | ngay nam ngoai phan dau sach").strip(" |")
    author, author_how = extract_author(text)
    best = dates[0] if dates else None
    day_full = best if (best and best["day_known"]) else None
    month_only = best if (best and not best["day_known"]) else None
    return dict(
        nam=year, nam_cach=year_how, nam_tin_cay=year_conf, nam_bang_chung=year_ev,
        ngay_du=day_full, ngay_thang=month_only, ngay_khac=dates[:8], ngay_ghi_chu=dates_note,
        tac_gia=author, tac_gia_cach=author_how,
        nha_xuat_ban=extract_publisher(text))
