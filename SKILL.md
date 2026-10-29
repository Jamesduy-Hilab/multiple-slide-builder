---
name: multiple-slide-builder
description: Automatically generate professional PowerPoint (.pptx) slides using Addsum Tech's visual layout templates (company-intro, quarterly-review, startup-pitch, corporate-training).
---

# Multiple Slide Builder Agent & Skill

Agent/Skill chuyên nghiệp dùng để tự động tạo và biên tập bài thuyết trình PowerPoint (`.pptx`) chất lượng cao với đồ họa vector nguyên bản (100% editable), không vỡ layout, dựa trên 4 bộ thiết kế chuẩn mực của Addsum Tech.

---

## 🎨 Kho Template Sẵn Có (Templates Catalog)

Nằm trong thư mục `templates/addsum/`:

1. **`company-intro`** (`company-intro.pptx` - 16 slides):
   - Phù hợp nhất cho: Giới thiệu công ty, hồ sơ năng lực (Capability Deck), ma trận sản phẩm, hệ sinh thái giải pháp, timeline và bản đồ phát triển.
   - Phong cách: Dark theme công nghệ cao, layout hiện đại, typography đậm nét.

2. **`quarterly-review`** (`quarterly-review.pptx` - 8 slides):
   - Phù hợp nhất cho: Báo cáo kết quả kinh doanh quý (QBR), họp cổ đông, bảng chỉ số KPIs, lưới so sánh dữ liệu và biểu đồ tăng trưởng.
   - Phong cách: Tinh gọn, nhấn mạnh số liệu và bảng biểu trực quan.

3. **`startup-pitch`** (`startup-pitch.pptx` - 12 slides):
   - Phù hợp nhất cho: Pitch deck gọi vốn (Seed, Series A), bài thuyết trình trước nhà đầu tư, mô hình kinh doanh, quy mô thị trường (TAM/SAM/SOM), Unit Economics.
   - Phong cách: Nổi bật, súc tích, thuyết phục mạnh mẽ.

4. **`corporate-training`** (`corporate-training.pptx` - 10 slides):
   - Phù hợp nhất cho: Giáo trình đào tạo nội bộ, workshop kỹ thuật, quy trình thao tác chuẩn (SOP), hướng dẫn từng bước.
   - Phong cách: Khối hộp thông tin mạch lạc, dễ tiếp thu và ghi nhớ.

---

## 🚀 Cách Thức Hoạt Động

### Cơ chế Vector & XML Manipulation:
1. **Slide Cloning:** Nhân bản chính xác cấu trúc slide XML từ template gốc, giữ nguyên các drawing shapes, connectors, layouts và relationships.
2. **Text Replacement:** Tìm kiếm và thay thế văn bản ở cấp độ `Run` và `Paragraph` để bảo toàn font chữ, cỡ chữ, màu sắc và độ đậm nhạt gốc.
3. **Logo Swapping:** Tự động phát hiện SHA1 part của logo mặc định trong layout/slide/master để thay thế bằng logo thương hiệu của bạn (`hilab-white.png`).
4. **Theme Recolor (Tùy chọn):** Đổi màu nhanh từ nền tối sang nền sáng thanh lịch với màu nhấn chủ đạo thương hiệu (Brand Accent).

---

## 💻 Hướng Dẫn Sử Dụng CLI

### Bước 1: Cài đặt thư viện phụ thuộc
```bash
pip install -r requirements.txt
```

### Bước 2: Tạo File Cấu hình Slide (`slides_config.json`)
Xác định template bạn muốn dùng và danh sách các slide cần clone cùng với từ khóa cần thay thế:

```json
{
  "template": "company-intro",
  "slides": [
    {
      "source_index": 0,
      "replacements": {
        "AN NVIDIA COMPANY & PRODUCT OVERVIEW": "HILAB TECHNOLOGY · CAPABILITY OVERVIEW",
        "The accelerated": "AI & CLOUD",
        "computing company.": "SOLUTIONS.",
        "AS OF JULY 2026   ·   EVERY FIGURE WEB-VERIFIED AGAINST NVIDIA FILINGS, PRESS RELEASES & INDUSTRY RESEARCH": "HILAB TECHNOLOGY  ·  PROJECT PROPOSAL"
      }
    },
    {
      "source_index": 2,
      "replacements": {
        "WHAT NVIDIA SELLS": "HỆ SINH THÁI GIẢI PHÁP HILAB",
        "Not a chip — the whole computing stack": "Giải pháp công nghệ toàn diện từ hạ tầng đến ứng dụng"
      }
    }
  ]
}
```

### Bước 3: Chạy Script Sinh Slide
```bash
python3 scripts/generate_addsum_slides.py \
  --template company-intro \
  --config examples/slides_config_example.json \
  --output output/my_deck.pptx
```

*Tùy chọn logo riêng:*
```bash
python3 scripts/generate_addsum_slides.py \
  --template company-intro \
  --config examples/slides_config_example.json \
  --logo resources/logo/hilab-white.png \
  --output output/my_deck.pptx
```

### Bước 4 (Tùy chọn): Chuyển Đổi Màu Sang Nền Sáng
```bash
python3 scripts/recolor_presentation_hilab.py \
  --input output/my_deck.pptx \
  --output output/my_deck_light.pptx \
  --bg FFFFFF \
  --accent CE2029 \
  --text 1E1D1B
```

---

## 💡 Lưu Ý Quan Trọng Khi Thiết Kế
1. **Độ dài câu chữ (Length Consistency):** Cố gắng giữ độ dài chuỗi text thay thế tương đương với văn bản gốc trong template để tránh hiện tượng text tràn khung hoặc co chữ.
2. **Định dạng hoa/thường (Case Sensitivity):** Nếu tiêu đề template dùng chữ `IN HOA`, hãy viết chuỗi thay thế bằng chữ in hoa để giữ thẩm mỹ ban đầu.
3. **Nhân bản linh hoạt (Duplication & Reordering):** Bạn có thể dùng 1 layout nhiều lần trong slide deck bằng cách khai báo nhiều object cùng `source_index` trong file JSON.
