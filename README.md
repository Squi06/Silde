# 7UP · Brand Deck

Bộ slide thuyết trình 12 trang về thương hiệu **7UP**, thiết kế 16:9 theo bộ nhận diện xanh–lime–đỏ,
viết bằng tiếng Việt, **sẵn sàng import vào Canva**.

## Xem nhanh

```bash
cd deck && python3 -m http.server 8000
# landing: http://localhost:8000
# bộ slide: http://localhost:8000/7up-slides.html   (← → để chuyển, N xem ghi chú, P chế độ in)
```

## Tệp chính

| Tệp | Nội dung |
| --- | --- |
| `deck/7UP-Brand-Deck.pptx` | Bộ slide 12 trang, 16:9, có ghi chú thuyết trình – **dùng để import Canva** |
| `deck/7up-slides.html` | Bản HTML tự chứa (font nhúng), điều hướng bằng bàn phím, không cần mạng |
| `deck/index.html` | Trang mở đầu nhỏ, dẫn tới slide / file tải về |
| `deck/preview/*.png` | Ảnh từng slide + contact sheet 3×4 |
| `CANVA.md` | Hướng dẫn import vào Canva, bảng màu, font, cách sửa |
| `deck/NOI-DUNG-THUYET-TRINH.md` | Kịch bản nói từng slide + Q&A dự phòng |

## Nội dung 12 slide

1. Bìa – *Vị chanh trong trẻo. Tinh thần khác biệt.*
2. Nội dung (6 phần)
3. Tổng quan: 1929 · 140+ quốc gia · 0 caffeine · <3% thị phần Mỹ
4. Hành trình 95 năm (8 dấu mốc)
5. Định vị & USP – “Uncola”, 4 trụ cột
6. Danh mục sản phẩm (Original, Zero Sugar, Cherry, Free, Revive, lon HORECA)
7. Khách hàng mục tiêu – 3 persona + insight
8. Marketing – Uncola, Spot, Make 7 Up Yours, Feels Good To Be You
9. Thị trường & đối thủ – 7UP vs Sprite vs Starry vs nhãn riêng
10. SWOT
11. Khuyến nghị – 4 việc trong 18 tháng
12. Kết luận & Q&A

## Dựng lại / chỉnh sửa

```bash
pip install python-pptx pillow fonttools brotli lxml

cd deck
python3 build_pptx.py   # -> 7UP-Brand-Deck.pptx
python3 preview.py      # -> preview/slide-XX.png + contact-sheet.png
python3 build_html.py   # -> 7up-slides.html (đọc trực tiếp từ file .pptx)
```

- Toàn bộ chữ nằm trong `deck/content.py` (một nguồn duy nhất cho mọi bản xuất).
- Màu, font, helper bố cục: `deck/design.py`; nhịp slide: `deck/build_pptx.py`.
- `deck/preview.py` là renderer nhẹ (PIL) để kiểm tra bố cục khi không có PowerPoint;
  `deck/build_html.py` chuyển `.pptx` → HTML nên bản xem trước luôn khớp bản in.

## Ghi chú

- Font dùng: **Montserrat** (tiêu đề) + **Open Sans** (nội dung) – đều có sẵn trong Canva.
- Ảnh bìa: ảnh minh hoạ do AI tạo, cắt tròn bằng hình học (`prstGeom=ellipse`) nên vẫn thay được
  ảnh khác trong Canva/PowerPoint mà giữ khung tròn.
- Số liệu thị phần là ước tính tổng hợp từ Beverage Digest và báo chí ngành; nên kiểm tra lại nếu
  bài yêu cầu trích dẫn học thuật.
