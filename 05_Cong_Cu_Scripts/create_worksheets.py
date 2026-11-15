import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from create_docs_helper import (
    setup_document_styles, set_cell_background, set_cell_margins,
    set_table_borders, set_callout_box
)

def add_header_block(doc, group_num, site_vn, site_en):
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_top.autofit = False
    c_l, c_r = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_l.width = Inches(3.2); c_r.width = Inches(3.5)

    p_l = c_l.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_l.add_run("TRƯỜNG THCS: .....................................\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p_l.add_run("TỔ BỘ MÔN: TIẾNG ANH - LỚP 8"); r.font.bold = True; r.font.size = Pt(10)

    p_r = c_r.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_r.add_run("DỰ ÁN STEM DẠY HỌC SỐ\n"); r.font.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_r.add_run("NĂM HỌC 2026 - 2027"); r.font.size = Pt(9.5)

    # Title Box
    tbl_title = doc.add_table(rows=1, cols=1)
    tbl_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_title.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_title, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_title.rows[0].cells[0].paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    r = p_box.add_run(f"PHIẾU HỌC TẬP DỰ ÁN STEM – NHÓM {group_num}\n")
    r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_box.add_run(f"DI TÍCH: {site_vn.upper()}\n")
    r.font.bold = True; r.font.size = Pt(13.5); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p_box.add_run(f"English name: {site_en}   |   Trạm số {group_num} – Lớp 8...\n")
    r.font.italic = True; r.font.bold = True; r.font.size = Pt(10.5)
    r = p_box.add_run("Sản phẩm: 01 Kịch bản tiếng Anh + 01 Audio Guide 45-60s + 01 Mã QR + Bộ 5 ảnh xuất bản lên Web")
    r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

def add_roles_table(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 1: BẢNG PHÂN CÔNG NHIỆM VỤ THÀNH VIÊN TRONG NHÓM"); r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=6, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    headers = ["Vai trò trong nhóm", "Họ và tên học sinh", "Nhiệm vụ cụ thể phụ trách"]
    rows_data = [
        ("1. Nhóm trưởng (Leader)", "........................................................", "Quản lý tiến độ, phân công, nộp bài lên website, điều phối thử nghiệm."),
        ("2. Thư ký / Biên kịch (Scriptwriter)", "........................................................", "Chắp bút kịch bản tiếng Anh 5 bước (80-110 từ), tra từ điển, chỉnh ngữ pháp."),
        ("3. Người thuyết minh (Speaker)", "........................................................", "Luyện phát âm chuẩn, ngữ điệu truyền cảm, thu âm Audio Guide 45-60s."),
        ("4. Nhiếp ảnh gia (Photographer)", "........................................................", "Chụp & chọn lọc bộ 5 ảnh (1 ảnh chính diện + 4 ảnh chi tiết kiến trúc/nội thất)."),
        ("5. Kỹ thuật viên số (Tech/Web)", "........................................................", "Lọc tạp âm (CapCut/Audacity), tải file Drive, tạo mã QR, thiết kế Canva A4.")
    ]

    for c_i, h in enumerate(headers):
        cell = tbl.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 40, 40, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    tbl.rows[0].cells[0].width = Inches(2.0)
    tbl.rows[0].cells[1].width = Inches(2.2)
    tbl.rows[0].cells[2].width = Inches(2.5)

    for r_i, r_data in enumerate(rows_data):
        row = tbl.rows[r_i + 1]
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            cell.width = [Inches(2.0), Inches(2.2), Inches(2.5)][c_i]
            set_cell_margins(cell, 40, 40, 50, 50)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
            if c_i == 0: r.font.bold = True

def add_step1_photos(doc, site_name, main_hint, sub_hints):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 1: CHỤP & THU THẬP BỘ 5 ẢNH DI TÍCH (Photo Collection)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_after = Pt(2)
    p_desc.add_run("Nhóm chuẩn bị 5 ảnh chụp sắc nét (khổ ngang, rõ nét, dung lượng < 3MB) theo danh mục gợi ý sau:")
    p_desc.runs[0].font.size = Pt(9.5); p_desc.runs[0].font.italic = True

    tbl = doc.add_table(rows=6, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    headers = ["Loại ảnh", "Góc chụp gợi ý cho nhóm", "Trạng thái nhóm"]
    rows = [
        ("Ảnh 1: Ảnh chính diện\n(Ảnh bìa lớn)", main_hint, "[  ] Đã chụp & chọn"),
        ("Ảnh phụ 1 (Kiến trúc)", sub_hints[0], "[  ] Đã chụp & chọn"),
        ("Ảnh phụ 2 (Chi tiết)", sub_hints[1], "[  ] Đã chụp & chọn"),
        ("Ảnh phụ 3 (Hiện vật)", sub_hints[2], "[  ] Đã chụp & chọn"),
        ("Ảnh phụ 4 (Trải nghiệm)", sub_hints[3], "[  ] Đã chụp & chọn")
    ]

    for c_i, h in enumerate(headers):
        cell = tbl.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 40, 40, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9)

    tbl.rows[0].cells[0].width = Inches(1.8)
    tbl.rows[0].cells[1].width = Inches(3.6)
    tbl.rows[0].cells[2].width = Inches(1.3)

    for r_i, r_data in enumerate(rows):
        row = tbl.rows[r_i + 1]
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            cell.width = [Inches(1.8), Inches(3.6), Inches(1.3)][c_i]
            set_cell_margins(cell, 35, 35, 50, 50)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)
            if c_i == 0 or c_i == 2: r.font.bold = True

def add_step2_script(doc, vocab_list, grammar_sample, scaffold_template):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 2: XÂY DỰNG KỊCH BẢN TIẾNG ANH 5 BƯỚC (Scriptwriting - Chuẩn 80 - 110 từ)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Scaffolding box
    tbl_vocab = doc.add_table(rows=2, cols=1)
    tbl_vocab.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_vocab.rows[0].cells[0].width = Inches(6.7); tbl_vocab.rows[1].cells[0].width = Inches(6.7)
    set_table_borders(tbl_vocab, color="B0C4DE", sz="4")
    set_cell_background(tbl_vocab.rows[0].cells[0], "F4F7FA")
    set_cell_margins(tbl_vocab.rows[0].cells[0], 40, 40, 50, 50)
    set_cell_margins(tbl_vocab.rows[1].cells[0], 40, 40, 50, 50)

    p0 = tbl_vocab.rows[0].cells[0].paragraphs[0]
    r = p0.add_run("Ngân hàng Từ vựng & Cấu trúc Ngữ pháp gợi ý cho nhóm:\n"); r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p0.add_run(f"• Từ vựng trọng tâm: {vocab_list}\n"); r.font.size = Pt(9)
    r = p0.add_run(f"• Mẫu câu ngữ pháp lớp 8: {grammar_sample}"); r.font.size = Pt(9); r.font.italic = True

    p1 = tbl_vocab.rows[1].cells[0].paragraphs[0]
    r = p1.add_run("Khung kịch bản 5 bước gợi ý (Nhóm thảo luận điền từ vào chỗ trống):\n"); r.font.bold = True; r.font.size = Pt(9.5)
    r = p1.add_run(scaffold_template); r.font.size = Pt(9); r.font.italic = True

    p_write = doc.add_paragraph()
    p_write.paragraph_format.space_before = Pt(4); p_write.paragraph_format.space_after = Pt(2)
    r = p_write.add_run("Bản kịch bản chính thức của nhóm (Thư ký viết vào khung dưới đây – đếm đủ 80-110 từ):"); r.font.bold = True; r.font.size = Pt(9.5)

    tbl_draft = doc.add_table(rows=1, cols=1)
    tbl_draft.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_draft.rows[0].cells[0].width = Inches(6.7)
    set_table_borders(tbl_draft, color="B0C4DE", sz="6")
    set_cell_margins(tbl_draft.rows[0].cells[0], 60, 60, 60, 60)
    p_box = tbl_draft.rows[0].cells[0].paragraphs[0]
    r = p_box.add_run(
        "1. Hook: ...........................................................................................................................................................\n"
        "2. Location: .......................................................................................................................................................\n"
        "3. History/Architecture: ........................................................................................................................................\n"
        "4. Fun Fact: ........................................................................................................................................................\n"
        "5. Advice & Closing: ..........................................................................................................................................\n"
        "-> Tổng số từ đếm được: ................ từ (Mục tiêu chuẩn: 80 - 110 từ).  |  Người đọc thử: ........................................"
    )
    r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

def add_step3_to_5(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 3, 4 & 5: THU ÂM, TẠO MÃ QR, XUẤT BẢN WEBSITE & THỬ NGHIỆM"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=4, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    steps = [
        ("Bước 3: Thu âm đạt chuẩn âm học (S)", "• Thu âm trong phòng kín, kéo rèm để triệt tiêu tiếng vang (Echo).\n• Đặt điện thoại cách miệng 1 gang tay (15 cm), nói chếch góc 45 độ để tránh luồng hơi thở làm rè âm (/p/, /t/).\n• Bật chế độ Reduce Noise (Khử ồn) trên CapCut hoặc Voice Memos. Thời lượng: 45 - 60 giây."),
        ("Bước 4: Xuất bản lên Website (T & E)", "• Tải file audio lên Google Drive, chuyển quyền 'Bất kỳ ai có đường liên kết đều có thể xem'.\n• Mở website dự án (index.html), bấm vào trạm của nhóm, tải lên 1 ảnh chính + 4 ảnh phụ, dán link audio và kịch bản.\n• Hệ thống tự động tạo mã QR -> Nhóm tải mã QR về để dán lên bảng Canva in màu A4."),
        ("Bước 5: Thử nghiệm Gallery Walk & Đo lường (M & E)", "• Tại Tiết 2, dán bảng A4 tại trạm. Đón du khách quét mã, lắng nghe audio và làm trắc nghiệm.\n• Nhóm ghi nhận số liệu: Tỷ lệ quét thành công: ..........% (≥ 90%)  |  Tỷ lệ nghe hiểu: ..........% (≥ 80%)  |  Điểm âm thanh: .......... / 5.0 sao.\n• Điền Phiếu cải tiến: Chỉ ra 1 lỗi kỹ thuật và giải pháp khắc phục của nhóm."),
        ("Đánh giá sản phẩm của nhóm", "[  ] Đạt kịch bản 80-110 từ    [  ] Đạt Audio 45-60s trong trẻo    [  ] Đủ 1+4 ảnh sắc nét\n[  ] Đã xuất bản thành công lên Website    [  ] Đạt tỷ lệ quét QR ≥ 90%\nChữ ký Trưởng nhóm: .......................................  |  Điểm của GV: ....................")
    ]

    for idx, (title, content) in enumerate(steps):
        row = tbl.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 35, 35, 50, 50); set_cell_margins(c1, 35, 35, 50, 50)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r0 = p0.add_run(title); r0.font.bold = True; r0.font.size = Pt(9)
        p1 = c1.paragraphs[0]; r1 = p1.add_run(content); r1.font.size = Pt(8.5)

def build_all_worksheets(filename="Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx"):
    doc = docx.Document()
    setup_document_styles(doc)

    # 4 Groups Data
    groups = [
        {
            "num": 1,
            "site_vn": "Bưu điện Trung tâm Sài Gòn",
            "site_en": "Saigon Central Post Office",
            "main_hint": "Toàn cảnh mặt tiền màu vàng kem cổ điển, nhìn thấy mái vòm và chiếc đồng hồ lớn phía trên cổng.",
            "sub_hints": [
                "Mái vòm sắt uốn lượn hình vòm bánh mì bên trong sảnh chính.",
                "Dãy bốt điện thoại công cộng bằng gỗ cổ kính hai bên lối vào.",
                "Hai tấm bản đồ cổ vẽ tay thời kỳ đầu của viễn thông Nam Kỳ.",
                "Quầy bán bưu thiếp lưu niệm hoặc du khách quốc tế ngồi viết thư."
            ],
            "vocab": "arched ceiling, telephone booths, vintage maps, magnificent, French architecture, send postcards.",
            "grammar": "was built between 1886 and 1891; is famous for its arched ceiling; You should send a postcard...",
            "template": "1. Hook: Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station?\n2. Location: It is the Saigon Central Post Office, located in District 1, right next to Notre Dame Cathedral.\n3. History/Architecture: This famous monument was built between 1886 and 1891. Inside, you can see a magnificent arched ceiling and wooden telephone booths.\n4. Fun Fact: Here is a fun fact: Many visitors still come here every day to write and send traditional postcards to their friends worldwide!\n5. Advice: You should take nice photos and remember to keep quiet. Enjoy your visit in Saigon!"
        },
        {
            "num": 2,
            "site_vn": "Nhà thờ Đức Bà Sài Gòn",
            "site_en": "Notre Dame Cathedral of Saigon",
            "main_hint": "Toàn cảnh mặt tiền hai tháp chuông uy nghiêm chụp từ Công viên 30/4 hoặc Công trường Công xã Paris.",
            "sub_hints": [
                "Đỉnh tháp chuông cao gần 60m vươn lên bầu trời xanh.",
                "Cận cảnh những viên gạch đỏ Marseille không bám rêu mốc sau hơn 140 năm.",
                "Tượng Đức Mẹ Hòa Bình bằng đá cẩm thạch trắng phía trước vườn hoa.",
                "Hàng rào hoa cổ kính hoặc góc phố sách bên hông nhà thờ."
            ],
            "vocab": "bell towers, red bricks, imported from France, marble statue, peaceful atmosphere, symbol of the city.",
            "grammar": "was constructed from 1877 to 1880; was built with red bricks; You shouldn't make loud noise...",
            "template": "1. Hook: Welcome to Saigon! Can you see the two tall bell towers rising into the sky?\n2. Location: This is Notre Dame Cathedral, standing proudly in Paris Square in the city center.\n3. History/Architecture: It was built between 1877 and 1880 with special red bricks imported directly from France. The bricks still keep their bright color today without plastering!\n4. Fun Fact: A fascinating fact: The cathedral has six giant bronze bells that weigh nearly thirty tons!\n5. Advice: When visiting, you should dress politely and speak softly to preserve the peaceful atmosphere. Have a wonderful time!"
        },
        {
            "num": 3,
            "site_vn": "Chợ Bến Thành",
            "site_en": "Ben Thanh Market",
            "main_hint": "Toàn cảnh tháp đồng hồ bốn mặt ở Cửa Nam chợ Bến Thành – biểu tượng thương mại sôi động của thành phố.",
            "sub_hints": [
                "Bức phù điêu đắp nổi hình con bò tót hoặc con cá trên cổng chợ.",
                "Gian hàng bán đồ lưu niệm thủ công mỹ nghệ, nón lá, túi thổ cẩm.",
                "Khu ẩm thực náo nhiệt với các món ăn đặc sản truyền thống (chè, bánh xèo, cà phê).",
                "Tiểu thương đang trò chuyện, giao tiếp tiếng Anh vui vẻ với du khách nước ngoài."
            ],
            "vocab": "clock tower, bustling market, traditional handicrafts, local food, souvenirs, practice bargaining.",
            "grammar": "was opened in 1914; is known as the most vibrant market; You should try authentic street food...",
            "template": "1. Hook: Are you looking for the most lively shopping place in Ho Chi Minh City? Let’s visit Ben Thanh Market!\n2. Location: Located right in the heart of District 1, it connects four major bustling streets.\n3. History/Architecture: Opened in 1914, its iconic south clock tower has become a world-famous symbol of Saigon. Inside, you can find hundreds of stalls selling silk, coffee, and crafts.\n4. Fun Fact: Here is a fun fact: You can taste almost all traditional southern Vietnamese street dishes in just one food section!\n5. Advice: You should try friendly bargaining with local sellers and buy some nice souvenirs for your family. Enjoy your shopping!"
        },
        {
            "num": 4,
            "site_vn": "Bến Nhà Rồng - Bảo tàng Hồ Chí Minh",
            "site_en": "Dragon Wharf - Ho Chi Minh Museum",
            "main_hint": "Toàn cảnh tòa nhà cổ Bến Nhà Rồng nhìn từ bờ sông Sài Gòn hoặc từ cầu Khánh Hội.",
            "sub_hints": [
                "Đôi rồng gốm tráng men xanh lớn trên nóc tòa nhà uốn lượn 'Lưỡng long chầu nguyệt'.",
                "Tượng đài hoặc bia tưởng niệm sự kiện Bác Hồ ra đi tìm đường cứu nước ngày 05/06/1911.",
                "Phòng trưng bày hiện vật lịch sử quý giá (chiếc vali, tài liệu, hiện vật cách mạng).",
                "Góc bến cảng nhìn ra sông Sài Gòn lộng gió với những con tàu tấp nập."
            ],
            "vocab": "ceramic dragons, historic wharf, patriot, set sail, national independence, precious artifacts.",
            "grammar": "was built in 1863; On June 5th, 1911, President Ho Chi Minh set sail...; You should listen respectfully...",
            "template": "1. Hook: Welcome to a historic landmark along the peaceful Saigon River: Dragon Wharf!\n2. Location: It is located in District 4, right at the meeting point of the river and the city port.\n3. History/Architecture: Originally built in 1863 as a French port house, the roof features two magnificent ceramic dragons facing the moon.\n4. Fun Fact: It holds great historical value: On June 5th, 1911, young patriot Nguyen Tat Thanh set sail abroad from here to find freedom for Vietnam!\n5. Advice: Today it is a museum preserving precious artifacts. You should visit respectfully and learn about our nation’s history. Thank you!"
        }
    ]

    for idx, grp in enumerate(groups):
        if idx > 0:
            doc.add_page_break()
        add_header_block(doc, grp["num"], grp["site_vn"], grp["site_en"])
        add_roles_table(doc)
        add_step1_photos(doc, grp["site_vn"], grp["main_hint"], grp["sub_hints"])
        add_step2_script(doc, grp["vocab"], grp["grammar"], grp["template"])
        add_step3_to_5(doc)

    doc.save(filename)
    print(f"File {filename} created successfully.")

if __name__ == "__main__":
    build_all_worksheets()
