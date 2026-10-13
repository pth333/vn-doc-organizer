# 📚 Công cụ sắp xếp tài liệu theo năm & đặt tên theo QĐ 68-QĐ/TW

Công cụ Python tự động **quét, đọc nội dung, phân loại theo năm** và **đổi tên** hàng loạt file `.pdf`, `.docx`, `.doc` theo quy định đặt tên của **QĐ 68-QĐ/TW**.

## Mục lục

- [Tính năng chính](#tính-năng-chính)
- [Chạy nhanh](#chạy-nhanh)
- [Tham số dòng lệnh](#tham-số-dòng-lệnh)
- [Cấu trúc tên file mới](#cấu-trúc-tên-file-mới)
- [File Excel báo cáo](#file-excel-báo-cáo)
- [Cách xác định năm và ngày](#cách-xác-định-năm-và-ngày)
- [Cấu hình `config.json`](#cấu-hình-configjson)
- [Cấu trúc thư mục dự án](#cấu-trúc-thư-mục-dự-án)
- [Ghi chú kỹ thuật & hạn chế](#ghi-chú-kỹ-thuật--hạn-chế)

---

## Tính năng chính

1. Quét **toàn bộ** file `.pdf`, `.docx`, `.doc` trong thư mục nguồn (**kể cả thư mục con**).
2. **Đọc nội dung** từng file để xác định **năm** (và ngày, tác giả nếu có).
3. **Copy** file vào thư mục đích, chia theo năm: `Sach_SAPXEP\2007\...`, `Sach_SAPXEP\2008\...`
   → file gốc **giữ nguyên**, không bị sửa hay xoá.
4. **Đổi tên** theo cấu trúc QĐ 68 (xem [bên dưới](#cấu-trúc-tên-file-mới)).
5. Ghi **1 file Excel báo cáo** vào thư mục đích.

---

## Chạy nhanh

```powershell
# Chạy theo config.json
python H:\tool_files\sap_xep_tai_lieu.py

# Xem trước, KHÔNG copy file nào
python H:\tool_files\sap_xep_tai_lieu.py --dry-run -v

# Thử nhanh 5 file đầu
python H:\tool_files\sap_xep_tai_lieu.py --dry-run --limit 5 -v

# Chỉ định thư mục khác
python H:\tool_files\sap_xep_tai_lieu.py --input "D:\Sach\Nguon" --output "D:\Sach\Da_sap_xep"

# Di chuyển file thay vì copy (cẩn thận, không undo được)
python H:\tool_files\sap_xep_tai_lieu.py --move
```

> 💡 **Khuyến nghị:** luôn chạy `--dry-run -v` trước để kiểm tra kết quả.
> Có thể chạy lại nhiều lần: file đã có ở thư mục đích sẽ bị bỏ qua và ghi chú trong cột `Canh bao`, trừ khi bật `"ghi_de_file_dich": true` trong config.

---

## Tham số dòng lệnh

| Tham số | Ý nghĩa |
|---------|---------|
| `--config` | Đường dẫn file cấu hình (mặc định `H:\tool_files\config.json`) |
| `--input` / `--output` | Ghi đè thư mục nguồn / đích trong config |
| `--dry-run` | Chỉ xem trước, không copy, không ghi Excel |
| `-v` | In từng file: tên cũ → tên mới, năm, tác giả |
| `--limit N` | Chỉ xử lý N file đầu (để thử) |
| `--move` | Di chuyển file (mặc định là copy, giữ file gốc) |

---

## Cấu trúc tên file mới

```
<TênChính><DDmmYY><Loại><ĐơnVị><M_T><SốLầnXB>.<ext>
ví dụ:  ho_chi_minh_bien_nien_tieu_su090523R04T02.PDF
```

| Nhóm | Ý nghĩa | Nguồn |
|------|---------|-------|
| 1 | Tên chính (không dấu, nối bằng `_`) | Tên file cũ |
| 2 | `DDmmYY` – thời gian tạo file | Xem [mục xác định ngày](#ngày-ddmmyy-nhóm-2--thời-gian-tạo-file) |
| 3 | `R` file ruột, `B` bìa, `D` sách điện tử, `K` văn bản khác | Tên file cũ, nếu không có thì đoán theo từ khoá |
| 4 | 2 số – đơn vị ban hành (01–21) | Tên file cũ, nếu không có thì đoán theo từ khoá |
| 5 | `M` xuất bản mới, `T` tái bản | Tên file cũ, nếu không có thì đoán theo từ khoá |
| 6 | 2 số – số lần xuất bản (`00` = lần đầu) | Tên file cũ, nếu không có thì đoán theo từ khoá |

---

## File Excel báo cáo

Tên file: `Bao_cao_sap_xep_<ngày giờ>.xlsx` (nằm trong thư mục đích).

| Sheet | Nội dung |
|-------|----------|
| `Da_xu_ly` | Mọi file đã xử lý: tên gốc, **tên mới**, năm, **tác giả**, ngày nhóm 2 + nguồn ngày, các nhóm 3–6, ngày sửa file, kích thước, đường dẫn gốc/mới, **cảnh báo** |
| `Thieu_nam_trong_noi_dung` | File **không đọc được năm** trong nội dung: tên file, đường dẫn, **ngày modified**, năm đọc được từ tên file cũ |
| `Thieu_ngay_trong_noi_dung` | File không đọc được ngày trong nội dung (đã dùng nguồn khác – ghi rõ nguồn) |
| `Loi_doc_file` | File **lỗi/không đọc được** (PDF hỏng, PDF scan ảnh, `.doc` hỏng…) kèm chi tiết lỗi |
| `Khong_ho_tro` | File khác định dạng (`.txt`, `.rar`…) và file tạm `~$...` của Word |
| `Tong_hop_theo_nam` | Số file, dung lượng, số file độ tin cậy thấp, số file lệch năm theo từng năm |
| `Cau_hinh` | Toàn bộ tham số đã dùng khi chạy |

> 👀 Cột **`Canh bao`** trong sheet `Da_xu_ly` là nơi cần xem **đầu tiên**: ghi rõ file nào bị đoán metadata, lệch năm, hoặc phải dùng ngày sửa file.

---

## Cách xác định năm và ngày

### Năm (dùng để chia thư mục)

Luôn đọc từ **nội dung file**, theo thứ tự ưu tiên:

| # | Dấu hiệu | Độ tin cậy |
|---|----------|-----------|
| 1 | Dòng bìa sách: `Hà Nội - 2007` | **Cao** |
| 2 | `xuất bản năm 2007`, `năm 2007 Nhà xuất bản ... xuất bản` | **Cao** |
| 3 | Năm đứng gần cụm `Nhà xuất bản` / `xuất bản` | Vừa |
| 4 | `tháng 3 năm 2008` ở phần đầu sách | Vừa |
| 5 | Năm xuất hiện nhiều nhất trong phần đầu sách, rồi trong toàn văn | Thấp |

Nếu **không đọc được năm nào**: file được copy **nguyên tên** vào thư mục `_KHONG_XAC_DINH_NAM` và ghi vào sheet `Thieu_nam_trong_noi_dung` kèm ngày modified.

### Ngày `DDmmYY` (nhóm 2 = "thời gian tạo file")

Thứ tự ưu tiên mặc định trong `config.json`:

1. `ten_file_cu` – 6 số có sẵn trong tên file cũ (nếu hợp lệ) ← **ưu tiên số 1**
2. `noi_dung_du_ngay` – ngày đủ ngày/tháng/năm đọc trong nội dung
3. `noi_dung_thang_nam` – chỉ có tháng/năm trong nội dung (lấy ngày = `01`)
4. `ngay_sua_file` – ngày chỉnh sửa file trên ổ đĩa

**Quy tắc thêm:**
- Ưu tiên ứng viên có **năm trùng năm của thư mục**, để tên file và thư mục thống nhất.
- Nếu tên file cũ có ngày thật nhưng khác năm xuất bản, ngày đó **được giữ nguyên** và có cảnh báo.

#### Đổi sang ưu tiên đọc ngày trong nội dung

Mở `config.json`, sửa:

```json
"ngay_uu_tien": ["noi_dung_du_ngay", "ten_file_cu", "noi_dung_thang_nam", "ngay_sua_file"]
```

> ⚠️ Trong sách/văn bản pháp luật, ngày trong nội dung thường là ngày **dẫn chứng tài liệu** (ví dụ sách năm 2007 nhưng trích dẫn văn bản 1949), nên có thể lệch xa ngày tạo file thật. Vì vậy mặc định công cụ ưu tiên ngày có trong tên file cũ.

---

## Cấu hình `config.json`

Các mục hay dùng:

```json
"thu_muc_nguon": "H:\\tool_files\\2007 file sách hoàn thành",
"thu_muc_dich":  "H:\\tool_files\\Sach_SAPXEP",
"dinh_dang_ho_tro": [".pdf", ".docx", ".doc"],
"copy_giu_file_goc": true,          // false = di chuyển file
"ngay_phai_cung_nam_voi_thu_muc": true,
"ten": {
  "viet_thuong": false,             // true = đổi tên chính thành chữ thường
  "bo_dau": true,                   // bỏ dấu tiếng Việt
  "giu_ky_tu": "-()",               // các ký tự được giữ lại trong tên chính
  "do_dai_toi_da": 170,
  "viet_hoa_phan_mo_rong": false    // true = đuôi file thành .PDF, .DOC
},
"mac_dinh": { "loai_file": "R", "don_vi": "04", "tinh_trang": "M", "so_lan_xuat_ban": "00" }
```

`doan_don_vi`, `doan_loai_file`, `doan_tinh_trang`: bảng **từ khoá → mã**, dùng khi tên file cũ không có sẵn mã. Sửa/thêm từ khoá theo thực tế của bạn.

---

## Cấu trúc thư mục dự án

```
H:\tool_files\
├── sap_xep_tai_lieu.py          # script chính
├── config.json                  # cấu hình
├── README.md                    # tài liệu này
├── _libs\                       # thư viện Python (pypdf, openpyxl, olefile)
├── Sach_SAPXEP\                 # KẾT QUẢ: thư mục theo năm + file Excel báo cáo
└── tools\
    ├── text_extract.py          # đọc text từ .pdf / .docx / .doc
    ├── doc_reader.py            # đọc .doc (Word 97-2003) không cần MS Word
    ├── tcvn3.py, tcvn3_table.py # giải mã văn bản TCVN3 (ABC) → Unicode
    ├── meta_extract.py          # tìm năm, ngày, tác giả trong nội dung
    ├── naming.py                # chuẩn hoá tên + ghép tên theo QĐ 68
    ├── download_libs.py         # tải thư viện (khi máy khác chưa có)
    ├── survey.py                # xem nhanh kết quả trích xuất toàn bộ file
    └── tao_du_lieu_test.py      # sinh dữ liệu test (file lỗi, file thiếu năm...)
```

---

## Ghi chú kỹ thuật & hạn chế

**Đã xử lý sẵn:**

- **File `.doc` cũ (Word 97-2003):** đọc trực tiếp bằng Python (định dạng OLE + piece table), **không cần cài Microsoft Word**, chạy nhanh và không bị treo.
- **Bảng mã TCVN3:** nhiều sách cũ lưu bằng TCVN3 (ABC) nên đọc ra lỗi kiểu `Hµ Néi`. Công cụ tự phát hiện và giải mã về Unicode (đã kiểm thử 10/10 mẫu).
- **File tạm và file bị chặn:** file `~$...` của Word được bỏ qua và ghi vào sheet `Khong_ho_tro`. File tải từ Internet có thể bị Windows chặn (Protected View) khiến Word treo – công cụ này không dùng Word nên không gặp vấn đề đó.
- **Tên trùng nhau:** file thứ hai được thêm `_2` và có cảnh báo.
- **File không đọc được năm:** vẫn được copy sang `_KHONG_XAC_DINH_NAM` nên không mất dữ liệu.

**Hạn chế cần lưu ý:**

- **PDF dạng ảnh scan** không có lớp text: ghi vào sheet `Loi_doc_file` (cần **OCR thủ công**).
- **PDF lỗi font cũ:** một số PDF chuyển từ font TCVN3 bị sai dấu (ví dụ "Chớnh" thay vì "Chính"). Văn bản vẫn đọc được để lấy năm, nhưng **tên tác giả có thể sai** – cần kiểm tra lại.