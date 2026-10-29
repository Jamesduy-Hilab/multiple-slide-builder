# 📊 Multiple Slide Builder Agent & CLI

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![Library](https://img.shields.io/badge/Library-python--pptx-orange.svg)](https://python-pptx.readthedocs.io/)
[![Agent Ready](https://img.shields.io/badge/AI%20Agent-Antigravity%20%7C%20Claude%20Code-purple.svg)](https://github.com/Jamesduy-Hilab)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Agent tự động hóa tạo slide thuyết trình PowerPoint chuyên nghiệp (.pptx) chuẩn đồ họa vector, giữ nguyên 100% khả năng chỉnh sửa (editable) dựa trên 4 bộ template hiện đại của Addsum Tech.**

---

## 🌟 Điểm Nổi Bật (Key Features)

- **100% Editable Vector Slides:** Không convert slide thành ảnh phẳng (raster image). Mọi shape, text box, table, icon và layout đều có thể click vào và chỉnh sửa trực tiếp trên PowerPoint, Google Slides hoặc Keynote.
- **Bảo toàn định dạng (Formatting Preservation):** Cơ chế thay thế văn bản ở cấp độ `Run` & `Paragraph` giúp giữ nguyên toàn bộ font chữ (Inter, Helvetica...), size chữ, màu sắc và thuộc tính in đậm/in nghiêng của thiết kế gốc.
- **Deep XML Slide Cloning:** Nhân bản các slide phức tạp mà không gây xung đột quan hệ (Relationship ID clashing) hay lỗi file PowerPoint.
- **Tự động thay thế Logo (Logo Swapping):** Tự động quét và thay thế logo mặc định bằng logo thương hiệu của bạn (`hilab-white.png`).
- **Hỗ trợ đổi tông màu thương hiệu (Brand Recolor):** Tích hợp công cụ chuyển đổi nhanh từ giao diện Dark Mode sang Light Mode với màu nhấn chủ đạo thương hiệu (Brand Accent).
- **Sẵn sàng tích hợp AI Agent (Skill Ready):** Tương thích hoàn toàn với Google Antigravity, Claude Code, Cursor, Windsurf dưới dạng Agent Skill tự hành.

---

## 📁 Cấu Trúc Thư Mục (Repository Structure)

```text
multiple-slide-builder/
├── SKILL.md                          # Hướng dẫn và định nghĩa Skill dành cho AI Agent
├── README.md                         # Tài liệu hướng dẫn sử dụng chi tiết
├── LICENSE                           # Giấy phép phần mềm MIT
├── requirements.txt                  # Danh sách thư viện Python phụ thuộc
├── .gitignore                        # Danh sách file loại trừ khi commit Git
├── examples/
│   └── slides_config_example.json    # File mẫu cấu hình slide JSON
├── resources/
│   └── logo/
│       ├── hilab-white.png           # Logo trắng hiển thị trên nền tối
│       └── hilab-red-wide.png        # Logo đỏ hiển thị trên nền sáng
├── scripts/
│   ├── generate_addsum_slides.py     # Script chính nhân bản và thay thế nội dung slide
│   └── recolor_presentation_hilab.py # Script đổi tông màu sang nền sáng và màu nhấn
└── templates/
    └── addsum/
        ├── company-intro.pptx        # 16 slides - Giới thiệu công ty & năng lực
        ├── corporate-training.pptx   # 10 slides - Đào tạo nội bộ, SOP, workshop
        ├── quarterly-review.pptx     # 8 slides  - Báo cáo quý, KPIs & dữ liệu
        └── startup-pitch.pptx        # 12 slides - Pitch deck gọi vốn nhà đầu tư
```

---

## 🎨 Danh Mục Template (Available Templates)

| Template | Số Slide | Mục đích tối ưu | Phong cách |
| :--- | :---: | :--- | :--- |
| **`company-intro`** | 16 | Hồ sơ năng lực, hệ sinh thái giải pháp, ma trận sản phẩm | Dark theme công nghệ cao, layout hiện đại, typography đậm |
| **`quarterly-review`** | 8 | Báo cáo QBR, họp cổ đông, bảng chỉ số KPIs, lưới số liệu | Tinh gọn, nhấn mạnh số liệu và bảng biểu trực quan |
| **`startup-pitch`** | 12 | Pitch deck gọi vốn (Seed/Series A), mô hình kinh doanh | Nổi bật, súc tích, bố cục thuyết phục nhà đầu tư |
| **`corporate-training`** | 10 | Giáo trình đào tạo nội bộ, quy trình chuẩn (SOP), workshop | Khối hộp thông tin mạch lạc, dễ tiếp thu và ghi nhớ |

---

## ⚙️ Cài Đặt & Sử Dụng (Quickstart)

### 1. Yêu cầu hệ thống
- Python 3.8+
- Thư viện `python-pptx`

### 2. Cài đặt thư viện
```bash
git clone https://github.com/Jamesduy-Hilab/multiple-slide-builder.git
cd multiple-slide-builder
pip install -r requirements.txt
```

---

## 🛠️ Hướng Dẫn Sử Dụng Chi Tiết

### Bước 1: Chuẩn bị file cấu hình JSON (`slides_config.json`)
Chỉ định template và mảng các slide cần nhân bản. Mỗi slide gồm `source_index` (vị trí slide gốc, tính từ 0) và cặp `target: replacement`:

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
        "From the graphics chip to the AI factory — NVIDIA's full-stack platform, product lines, business model, and key figures at a glance.": "Kiến trúc hệ thống, giải pháp AI/Data chuyên sâu và năng lực công nghệ tùy biến toàn diện cho doanh nghiệp.",
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

### Bước 2: Chạy lệnh sinh slide
```bash
python3 scripts/generate_addsum_slides.py \
  --template company-intro \
  --config examples/slides_config_example.json \
  --output output/my_deck.pptx
```

*Muốn truyền logo thương hiệu tùy chỉnh:*
```bash
python3 scripts/generate_addsum_slides.py \
  --template company-intro \
  --config examples/slides_config_example.json \
  --logo resources/logo/hilab-white.png \
  --output output/my_deck.pptx
```

### Bước 3: Đổi màu sang nền sáng chuẩn thương hiệu (Tùy chọn)
Nếu bạn muốn bộ slide có nền trắng thanh lịch với màu nhấn thương hiệu (ví dụ đỏ `#CE2029`):
```bash
python3 scripts/recolor_presentation_hilab.py \
  --input output/my_deck.pptx \
  --output output/my_deck_light.pptx \
  --bg FFFFFF \
  --accent CE2029 \
  --text 1E1D1B
```

---

## 🤖 Tích Hợp Vào AI Agent (Antigravity & Claude Code)

Repository này được đóng gói theo chuẩn Agent Skill.

### Cách 1: Cài đặt Toàn cục (Global Agent Skill)
1. Copy thư mục `multiple-slide-builder` vào thư mục cấu hình Agent trên máy của bạn:
   - macOS / Linux: `~/.gemini/config/skills/` hoặc `~/.claude/skills/`
2. Khởi động lại IDE / Agent CLI.
3. Ra lệnh tự nhiên cho Agent:
   - *"Tạo slide giới thiệu năng lực công ty bằng multiple-slide-builder theo template company-intro"*
   - *"Nhân bản slide số 0 và số 2 của template startup-pitch và điền thông tin dự án X vào"*

### Cách 2: Sử dụng trong từng Workspace / Dự án (Project-Scoped)
1. Đặt thư mục này vào `.agents/skills/multiple-slide-builder/` bên trong workspace dự án.
2. AI Assistant sẽ tự động nhận diện và đọc file `SKILL.md` để kích hoạt bộ công cụ.

---

## 📄 License

Dự án được phân phối dưới giấy phép **MIT License**. Xem chi tiết tại file [LICENSE](LICENSE).
