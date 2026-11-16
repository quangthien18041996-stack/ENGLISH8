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
    r = p_r.add_run("HỘI THI THIẾT KẾ CHỦ ĐỀ STEM & DẠY HỌC SỐ\n"); r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_r.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO TP. HỒ CHÍ MINH"); r.font.size = Pt(9); r.font.italic = True

    # Title Box
    tbl_title = doc.add_table(rows=1, cols=1)
    tbl_title.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_title.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_title, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_title.rows[0].cells[0].paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    r = p_box.add_run(f"PHIẾU HỌC TẬP DỰ ÁN STEM TIẾNG ANH 8 – NHÓM {group_num}\n")
    r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_box.add_run(f"TRẠM DI TÍCH: {site_vn.upper()}\n")
    r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p_box.add_run(f"International Name: {site_en}   |   Trạm số {group_num} – Lớp 8...\n")
    r.font.italic = True; r.font.bold = True; r.font.size = Pt(10)
    r = p_box.add_run("Sản phẩm số: 01 Kịch bản 5 bước (80-110 từ) + 01 Audio Guide MP3 (45-60s) + Bộ 5 ảnh kiến trúc + Mã QR tương tác trực tiếp trên Cổng học liệu số (https://english8-two.vercel.app/)")
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

def add_roles_table(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 1: PHÂN CÔNG NHIỆM VỤ CÁC THÀNH VIÊN TRONG NHÓM (5 VAI TRÒ CHUẨN STEM)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=6, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    headers = ["Vai trò chuyên trách", "Họ và tên học sinh", "Nhiệm vụ cụ thể"]
    roles = [
        ("1. Trưởng nhóm (Team Leader)", "........................................................", "Điều phối chung, quản lý tiến độ, kiểm soát số từ kịch bản và chất lượng âm thanh."),
        ("2. Biên tập kịch bản (Scriptwriter)", "........................................................", "Soạn thảo kịch bản 5 bước tiếng Anh (80-110 từ), tra cứu từ vựng IPA, dịch nghĩa."),
        ("3. Phát thanh viên (Voice Actor)", "........................................................", "Luyện phát âm chuẩn âm đuôi, ngữ điệu truyền cảm, thu âm Audio MP3 (45-60 giây)."),
        ("4. Kỹ thuật số & QR (Tech Specialist)", "........................................................", "Khử ồn âm thanh, nộp bài lên Cổng học liệu số, tạo mã QR, theo dõi lượt thả tim & bình luận."),
        ("5. Nhiếp ảnh & Thuyết trình (Presenter)", "........................................................", "Chụp 1 ảnh bìa + 4 ảnh chi tiết kiến trúc STEM; thuyết trình poster tại buổi triển lãm.")
    ]

    for c_i, h_txt in enumerate(headers):
        cell = tbl.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 40, 40, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    tbl.rows[0].cells[0].width = Inches(2.2)
    tbl.rows[0].cells[1].width = Inches(2.1)
    tbl.rows[0].cells[2].width = Inches(2.4)

    for r_i, (r_role, r_name, r_task) in enumerate(roles):
        row = tbl.rows[r_i + 1]
        row.cells[0].width = Inches(2.2); row.cells[1].width = Inches(2.1); row.cells[2].width = Inches(2.4)
        for c_i, val in enumerate([r_role, r_name, r_task]):
            cell = row.cells[c_i]
            set_cell_margins(cell, 35, 35, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(8.5)
            if c_i == 0: r.font.bold = True

def add_step1_photos(doc, site_name, main_hint, sub_hints):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 1: KHẢO SÁT THỰC ĐỊA & CHỤP BỘ 5 ẢNH DI TÍCH KIẾN TRÚC (1 ẢNH BÌA + 4 GÓC CHỤP CHI TIẾT)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=6, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    headers = ["Loại ảnh yêu cầu", "Gợi ý góc chụp kỹ thuật (STEM)", "Ghi chú thực hiện của nhóm"]
    rows = [
        ("Ảnh 1: Ảnh bìa chính (Main Cover)", f"Toàn cảnh {site_name}. {main_hint}", "[  ] Đã chụp  |  Tên file: cover_main.jpg"),
        ("Ảnh 2: Chi tiết kiến trúc 1", sub_hints[0], "[  ] Đã chụp  |  Tên file: sub1_arch.jpg"),
        ("Ảnh 3: Chi tiết kiến trúc 2", sub_hints[1], "[  ] Đã chụp  |  Tên file: sub2_structure.jpg"),
        ("Ảnh 4: Hiện vật / Không gian văn hóa", sub_hints[2], "[  ] Đã chụp  |  Tên file: sub3_culture.jpg"),
        ("Ảnh 5: Hoạt động trải nghiệm / Con người", sub_hints[3], "[  ] Đã chụp  |  Tên file: sub4_activity.jpg")
    ]

    for c_i, h_txt in enumerate(headers):
        cell = tbl.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 40, 40, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    tbl.rows[0].cells[0].width = Inches(2.2)
    tbl.rows[0].cells[1].width = Inches(2.6)
    tbl.rows[0].cells[2].width = Inches(1.9)

    for r_i, r_data in enumerate(rows):
        row = tbl.rows[r_i + 1]
        row.cells[0].width = Inches(2.2); row.cells[1].width = Inches(2.6); row.cells[2].width = Inches(1.9)
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            set_cell_margins(cell, 35, 35, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(8.5)
            if c_i == 0: r.font.bold = True

def add_step2_script(doc, vocab_list, grammar_sample, scaffold_template):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 2: BIÊN SOẠN KỊCH BẢN TIẾNG ANH 5 BƯỚC (CHUẨN LỚP 8: 80 - 110 TỪ)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Scaffolding box
    tbl_vocab = doc.add_table(rows=2, cols=1)
    tbl_vocab.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_vocab.rows[0].cells[0].width = Inches(6.7); tbl_vocab.rows[1].cells[0].width = Inches(6.7)
    set_table_borders(tbl_vocab, color="B0C4DE", sz="4")
    set_cell_background(tbl_vocab.rows[0].cells[0], "F4F7FA")
    set_cell_margins(tbl_vocab.rows[0].cells[0], 35, 35, 45, 45)
    set_cell_margins(tbl_vocab.rows[1].cells[0], 35, 35, 45, 45)

    p0 = tbl_vocab.rows[0].cells[0].paragraphs[0]
    r = p0.add_run("Ngân hàng Từ vựng IPA & Cấu trúc Ngữ pháp Lớp 8 gợi ý cho nhóm:\n"); r.font.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p0.add_run(f"• Từ vựng trọng tâm: {vocab_list}\n"); r.font.size = Pt(8.5)
    r = p0.add_run(f"• Cấu trúc ngữ pháp Unit 5 & 11: {grammar_sample}"); r.font.size = Pt(8.5); r.font.italic = True

    p1 = tbl_vocab.rows[1].cells[0].paragraphs[0]
    r = p1.add_run("Khung kịch bản 5 bước mẫu (Nhóm thảo luận điền dữ liệu phù hợp):\n"); r.font.bold = True; r.font.size = Pt(9)
    r = p1.add_run(scaffold_template); r.font.size = Pt(8.5); r.font.italic = True

    p_write = doc.add_paragraph()
    p_write.paragraph_format.space_before = Pt(4); p_write.paragraph_format.space_after = Pt(2)
    r = p_write.add_run("Bản kịch bản chính thức của nhóm (Thư ký hoàn thiện – kiểm tra số từ đạt 80 - 110 từ):"); r.font.bold = True; r.font.size = Pt(9.5)

    tbl_draft = doc.add_table(rows=1, cols=1)
    tbl_draft.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_draft.rows[0].cells[0].width = Inches(6.7)
    set_table_borders(tbl_draft, color="B0C4DE", sz="6")
    set_cell_margins(tbl_draft.rows[0].cells[0], 50, 50, 50, 50)
    p_box = tbl_draft.rows[0].cells[0].paragraphs[0]
    r = p_box.add_run(
        "1. Hook (Mở đầu hấp dẫn): ........................................................................................................................................................................................\n"
        "2. Location (Vị trí địa lý): ........................................................................................................................................................................................\n"
        "3. History/Architecture (Lịch sử & Kiến trúc STEM): .........................................................................................................................................................\n"
        "4. Fun Fact (Điểm thú vị độc đáo): ................................................................................................................................................................................\n"
        "5. Advice & Closing (Lời khuyên & Chào tạm biệt): .............................................................................................................................................................\n"
        "-> Tổng số từ đếm được: .............. từ (Chuẩn yêu cầu: 80 - 110 từ).   |   Thời gian đọc thử: .............. giây (Chuẩn: 45 - 60s)."
    )
    r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

def add_step3_to_5(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("BƯỚC 3, 4 & 5: THU ÂM CHUẨN STUDIO, NỘP BÀI LÊN WEBSITE, XÁC THỰC DỮ LIỆU & BÁO CÁO"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=4, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    steps = [
        ("Bước 3: Thu âm đạt chuẩn âm học (Science & Tech)", "• Chọn không gian kín, kéo rèm hoặc phòng thu yên tĩnh để triệt tiêu tiếng vang (Echo), độ ồn nền < 45 dB.\n• Đặt micro cách miệng 15 cm, nói chếch góc 45 độ để tránh luồng hơi thở làm rè các phụ âm bật hơi (/p/, /t/, /k/).\n• Bật tính năng Khử ồn (Noise Reduction) trên CapCut hoặc Voice Memos. Thời lượng: 45 - 60 giây, phát âm rõ âm cuối (/s/, /z/, /ed/)."),
        ("Bước 4: Nộp bài lên Cổng học liệu số (Engineering & Tech)", "• Truy cập Cổng học liệu số: https://english8-two.vercel.app/ (hoặc https://quangthien18041996-stack.github.io/ENGLISH8/).\n• Nhấn nút 'Student Submit' (Học sinh nộp bài), tải lên 1 ảnh bìa + 4 ảnh kiến trúc chi tiết, dán kịch bản 5 bước và tải file MP3.\n• Hệ thống tự động đếm số từ chuẩn Lớp 8 và đưa bài vào hàng chờ kiểm duyệt: 'Chờ Giáo viên duyệt...'.\n• Giáo viên bộ môn mở tính năng 'Teacher Verification' (Giáo viên kiểm duyệt) để thẩm định chất lượng và tích chọn: 'Thông tin đã được phê duyệt ✓' theo chuẩn CV 5512.\n• Hệ thống tự động tạo Mã QR Audio Guide liên kết trang Infographic tương tác."),
        ("Bước 5: Báo cáo Thuyết trình & Đo lường thực tế (Math & Art)", "• Tại Tiết 2, nhóm trình chiếu Poster di sản trên máy chiếu lớp học bằng Chế độ thuyết trình (Presentation Mode).\n• Nhấn ô mã QR để phóng to mã QR khổ lớn (240x240px) trên máy chiếu để cả lớp quét từ xa bằng điện thoại.\n• Học sinh dưới lớp quét mã QR trên điện thoại: nghe audio, nhấn '❤️ Thả tim cổ vũ' (hệ thống Real-time Cloud Sync nhảy số tức thì lên màn hình máy chiếu), gửi bình luận nhận xét nhóm bạn (Peer Feedback) và làm 3 câu trắc nghiệm nhanh.\n• Nhóm ghi nhận số liệu: Tỷ lệ quét thành công: ........% (≥ 90%) | Số lượt tim cổ vũ: ........ tim | Tỷ lệ trả lời đúng trắc nghiệm: ........% (≥ 80%)."),
        ("Đánh giá sản phẩm của nhóm", "[  ] Kịch bản đạt 80-110 từ    [  ] Audio chuẩn 45-60s trong trẻo    [  ] Đủ 1+4 ảnh sắc nét\n[  ] Đã xuất bản lên Cổng học liệu số    [  ] Được Giáo viên cấp huy hiệu 'Đã phê duyệt ✓'\n[  ] Nhận tương tác thả tim & bình luận thời gian thực\nChữ ký Trưởng nhóm: .......................................  |  Điểm & Nhận xét của GV: ....................")
    ]

    for idx, (title, content) in enumerate(steps):
        row = tbl.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 35, 35, 45, 45); set_cell_margins(c1, 35, 35, 45, 45)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r0 = p0.add_run(title); r0.font.bold = True; r0.font.size = Pt(8.5)
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
            "main_hint": "Toàn cảnh mặt tiền màu vàng kem phong cách châu Âu, mái vòm cong lớn và chiếc đồng hồ cổ kính phía trên cổng chính.",
            "sub_hints": [
                "Mái vòm sắt uốn lượn chịu lực vắt ngang trần sảnh chính (Kiến trúc đường sắt châu Âu).",
                "Dãy bốt điện thoại công cộng bằng gỗ cổ kính hai bên lối vào.",
                "Hai bức họa bản đồ lịch sử quý giá vẽ tay: 'Saigon et ses environs 1892'.",
                "Quầy bán bưu thiếp lưu niệm hoặc khách quốc tế đang ngồi viết thư tay gửi khắp năm châu."
            ],
            "vocab": "arched ceiling (/ɑːtʃt ˈsiːlɪŋ/), telephone booths (/ˈtelɪfəʊn buːðz/), vintage maps (/ˈvɪntɪdʒ mæps/), colonial architecture (/kəˈləʊniəl ˈɑːkɪtektʃə/), handwritten postcard (/ˌhændˈrɪtn ˈpəʊstkɑːd/).",
            "grammar": "was built between 1886 and 1891 (Past Simple Passive); is famous for its arched ceiling; You should send a postcard...",
            "template": "1. Hook: Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station?\n2. Location: It is the Saigon Central Post Office, located in District 1, right next to Notre Dame Cathedral.\n3. History/Architecture: This famous monument was built between 1886 and 1891. Inside, you can see a magnificent arched ceiling and wooden telephone booths.\n4. Fun Fact: Here is a fun fact: Many visitors still come here every day to write and send traditional postcards to their friends worldwide!\n5. Advice: You should take nice photos and remember to keep quiet. Enjoy your visit in Saigon!"
        },
        {
            "num": 2,
            "site_vn": "Nhà thờ Đức Bà Sài Gòn",
            "site_en": "Notre Dame Cathedral of Saigon",
            "main_hint": "Toàn cảnh mặt tiền hai tháp chuông đôi 60.5m uy nghiêm chụp từ Công trường Công xã Paris hoặc Công viên 30/4.",
            "sub_hints": [
                "Đỉnh tháp chuông đôi cao vút vươn lên nền trời xanh mang phong cách Romanesque kết hợp Gothic.",
                "Cận cảnh những viên gạch đỏ Marseille không trát vữa nhưng giữ nguyên màu tươi hơn 140 năm.",
                "Tượng Đức Mẹ Hòa Bình bằng đá cẩm thạch trắng nguyên khối phía trước vườn hoa.",
                "Hệ thống 6 quả chuông đồng lớn nặng gần 30 tấn được điều khiển bằng hệ thống cơ khí chuẩn xác."
            ],
            "vocab": "bell towers (/bel ˈtaʊəz/), red bricks (/red brɪks/), imported from France (/ɪmˈpɔːtɪd frəm frɑːns/), marble statue (/ˈmɑːbl ˈstætʃuː/), peaceful atmosphere (/ˈpiːsfl ˈætməsfɪə/).",
            "grammar": "was constructed from 1877 to 1880; was built with red bricks; You shouldn't make loud noise; Remember to dress politely...",
            "template": "1. Hook: Welcome to Saigon! Can you see the two tall bell towers rising into the sky?\n2. Location: This is Notre Dame Cathedral, standing proudly in Paris Square in the city center.\n3. History/Architecture: It was built between 1877 and 1880 with special red bricks imported directly from France. The bricks still keep their bright color today without plastering!\n4. Fun Fact: A fascinating fact: The cathedral has six giant bronze bells that weigh nearly thirty tons!\n5. Advice: When visiting, you should dress politely and speak softly to preserve the peaceful atmosphere. Have a wonderful time!"
        },
        {
            "num": 3,
            "site_vn": "Chợ Bến Thành",
            "site_en": "Ben Thanh Market",
            "main_hint": "Toàn cảnh tháp đồng hồ 4 mặt ở Cửa Nam chợ Bến Thành – biểu tượng nhận diện văn hóa và giao thương sôi động bậc nhất của TP.HCM.",
            "sub_hints": [
                "Bức phù điêu đắp nổi tinh xảo hình con bò tót hoặc con cá trên cổng chợ.",
                "Gian hàng bày bán đồ thủ công mỹ nghệ, nón lá, áo dài, túi thổ cẩm truyền thống.",
                "Khu ẩm thực náo nhiệt với các món đặc sản Nam Bộ (bánh xèo, chè ba màu, cà phê vợt).",
                "Tiểu thương đang trò chuyện, giao tiếp tiếng Anh vui vẻ và hướng dẫn du khách trả giá thân thiện."
            ],
            "vocab": "clock tower (/klɒk ˈtaʊə/), bustling market (/ˈbʌslɪŋ ˈmɑːkɪt/), traditional handicrafts (/trəˈdɪʃənl ˈhændikrɑːfts/), local street food (/ˈləʊkl striːt fuːd/), bargain politely (/ˈbɑːɡən pəˈlaɪtli/).",
            "grammar": "was opened in 1914; is known as the most vibrant market; You should try authentic street food; Remember to practice bargaining...",
            "template": "1. Hook: Are you looking for the most lively shopping place in Ho Chi Minh City? Let’s visit Ben Thanh Market!\n2. Location: Located right in the heart of District 1, it connects four major bustling streets.\n3. History/Architecture: Opened in 1914, its iconic south clock tower has become a world-famous symbol of Saigon. Inside, you can find hundreds of stalls selling silk, coffee, and crafts.\n4. Fun Fact: Here is a fun fact: You can taste almost all traditional southern Vietnamese street dishes in just one food section!\n5. Advice: You should try friendly bargaining with local sellers and buy some nice souvenirs for your family. Enjoy your shopping!"
        },
        {
            "num": 4,
            "site_vn": "Bến Nhà Rồng - Bảo tàng Hồ Chí Minh",
            "site_en": "Dragon Wharf - Ho Chi Minh Museum",
            "main_hint": "Toàn cảnh tòa nhà cổ kính Bến Nhà Rồng nhìn từ bờ sông Sài Gòn hoặc từ cầu Khánh Hội với kiến trúc Pháp kết hợp Á Đông.",
            "sub_hints": [
                "Đôi rồng gốm tráng men xanh uốn lượn 'Lưỡng long chầu nguyệt' tinh xảo trên nóc tòa nhà.",
                "Bia tưởng niệm sự kiện người thanh niên yêu nước Nguyễn Tất Thành ra đi tìm đường cứu nước ngày 05/06/1911.",
                "Phòng trưng bày các hiện vật lịch sử cách mạng quý giá (chiếc vali mây, tài liệu, hiện vật bôn ba hải ngoại).",
                "Góc bến cảng lịch sử hướng ra sông Sài Gòn lộng gió với những con tàu quốc tế qua lại."
            ],
            "vocab": "ceramic dragons (/səˈræmɪk ˈdræɡənz/), historic wharf (/hɪˈstɒrɪk wɔːf/), patriot (/ˈpeɪtriət/), set sail (/set seɪl/), national independence (/ˈnæʃnəl ˌɪndɪˈpendəns/), precious artifacts (/ˈpreʃəs ˈɑːtɪfækts/).",
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
    import os
    target_dir = os.path.join("01_Ho_So_Du_Thi_SGDDT", "Folder 2 - Ke hoach bai day")
    os.makedirs(target_dir, exist_ok=True)
    target_file = os.path.join(target_dir, "Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx")
    build_all_worksheets(target_file)
