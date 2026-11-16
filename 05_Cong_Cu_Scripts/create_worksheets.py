import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from create_docs_helper import (
    setup_document_styles, set_cell_background, set_cell_margins,
    set_table_borders, set_callout_box
)

def add_header_block(doc, group_num, site_vn, site_en):
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_top.autofit = False
    c_l, c_r = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_l.width = Inches(3.2); c_r.width = Inches(3.5)

    p_l = c_l.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_l.add_run("TRƯỜNG THCS: ........................................\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p_l.add_run("LỚP: 8......  |  TỔ: .........."); r.font.bold = True; r.font.size = Pt(10)

    p_r = c_r.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_r.add_run("CHỦ ĐỀ STEM TIẾNG ANH 8\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p_r.add_run("DỰ ÁN: SAIGON HERITAGE QR AUDIO GUIDE"); r.font.italic = True; r.font.size = Pt(9.5)

    tbl_name = doc.add_table(rows=1, cols=1)
    tbl_name.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_name.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_name, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_name.rows[0].cells[0].paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_box.add_run(f"PHIẾU HỌC TẬP SỐ HÓA – NHÓM {group_num}\n")
    r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_box.add_run(f"ĐỊA DANH PHỤ TRÁCH: {site_vn.upper()} ({site_en.upper()})\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p_box.add_run("Cổng học liệu số: https://english8-two.vercel.app/")
    r.font.italic = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)

def add_roles_table(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 1: PHÂN CÔNG NHIỆM VỤ CÁC THÀNH VIÊN TRONG NHÓM (5 VAI TRÒ CHUẨN STEM)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl = doc.add_table(rows=6, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False
    set_table_borders(tbl, color="B0C4DE", sz="4")

    headers = ["Vai trò chuyên trách", "Họ và tên học sinh", "Nhiệm vụ cụ thể"]
    roles = [
        ("1. Trưởng nhóm (Team Leader)", "........................................................", "Điều phối chung, quản lý tiến độ, đếm số từ kịch bản (80-110 từ)."),
        ("2. Người tìm ảnh (Photo Researcher)", "........................................................", "Tìm và chọn lọc 3 - 5 bức ảnh chụp rõ nét, góc chụp đẹp về di tích."),
        ("3. Người viết bài (Scriptwriter)", "........................................................", "Soạn thảo bài giới thiệu tiếng Anh 80 - 110 từ theo cấu trúc 5 bước."),
        ("4. Phát thanh viên (Voice Actor)", "........................................................", "Luyện phát âm chuẩn âm đuôi (/s/, /z/, /ed/), thu âm audio 45 - 60 giây."),
        ("5. Người thuyết trình (Presenter)", "........................................................", "Chuẩn bị bài thuyết trình trước lớp ở Tiết 2 cùng tệp audio của nhóm.")
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

def add_step1_photos(doc, site_vn, main_hint, sub_hints):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"PHẦN 2: NHIỆM VỤ 1 – TÌM KIẾM VÀ CHỌN LỌC HÌNH ẢNH TIÊU BIỂU VỀ {site_vn.upper()}"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    
    p_desc = doc.add_paragraph()
    p_desc.add_run("Yêu cầu: Người tìm ảnh cùng nhóm chọn 01 ảnh toàn cảnh bao quát và 03 - 04 ảnh cận cảnh chi tiết kiến trúc đặc thù (ảnh rõ nét, ánh sáng đẹp, không bị mờ):\n")

    tbl_p = doc.add_table(rows=5, cols=3)
    tbl_p.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_p.autofit = False
    set_table_borders(tbl_p, color="B0C4DE", sz="4")

    p_hd = ["STT / Góc ảnh", "Gợi ý góc chụp / Nét kiến trúc cần tìm", "Ghi chú nguồn ảnh / Tên tệp ảnh của nhóm"]
    for c_i, h_txt in enumerate(p_hd):
        cell = tbl_p.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 35, 35, 45, 45)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9)

    tbl_p.rows[0].cells[0].width = Inches(1.5)
    tbl_p.rows[0].cells[1].width = Inches(3.2)
    tbl_p.rows[0].cells[2].width = Inches(2.0)

    rows_data = [
        ("Ảnh 1 (Bắt buộc)", f"Toàn cảnh bao quát: {main_hint}", "[ Tệp: .................................................... ]"),
        ("Ảnh 2 (Chi tiết 1)", f"Chi tiết đặc trưng 1: {sub_hints[0]}", "[ Tệp: .................................................... ]"),
        ("Ảnh 3 (Chi tiết 2)", f"Chi tiết đặc trưng 2: {sub_hints[1]}", "[ Tệp: .................................................... ]"),
        ("Ảnh 4 (Chi tiết 3)", f"Chi tiết đặc trưng 3: {sub_hints[2]}", "[ Tệp: .................................................... ]")
    ]

    for r_i, (c0_val, c1_val, c2_val) in enumerate(rows_data):
        row = tbl_p.rows[r_i + 1]
        row.cells[0].width = Inches(1.5); row.cells[1].width = Inches(3.2); row.cells[2].width = Inches(2.0)
        for c_i, val in enumerate([c0_val, c1_val, c2_val]):
            cell = row.cells[c_i]
            set_cell_margins(cell, 30, 30, 40, 40)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(8.5)
            if c_i == 0: r.font.bold = True

def add_step2_script(doc, vocab, grammar, template):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 3: NHIỆM VỤ 2 – SOẠN THẢO BÀI GIỚI THIỆU TIẾNG ANH (80 - 110 TỪ)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_v = doc.add_paragraph()
    r = p_v.add_run("• Ngân hàng từ vựng gợi ý (Vocabulary): "); r.font.bold = True; r.font.size = Pt(9)
    r = p_v.add_run(vocab); r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_g = doc.add_paragraph()
    r = p_g.add_run("• Cấu trúc ngữ pháp trọng tâm (Grammar Focus): "); r.font.bold = True; r.font.size = Pt(9)
    r = p_g.add_run(grammar); r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x00, 0x66, 0x00)

    p_t = doc.add_paragraph()
    r = p_t.add_run("• Kịch bản tham khảo chuẩn 5 bước:\n"); r.font.bold = True; r.font.size = Pt(9)
    r = p_t.add_run(template); r.font.size = Pt(8.5); r.font.italic = True

    p_draft = doc.add_paragraph()
    p_draft.paragraph_format.space_before = Pt(4); p_draft.paragraph_format.space_after = Pt(2)
    r = p_draft.add_run("BẢN THẢO BÀI GIỚI THIỆU CHÍNH THỨC CỦA NHÓM (Dung lượng: 80 - 110 từ):"); r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    tbl_draft = doc.add_table(rows=5, cols=1)
    tbl_draft.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_draft.autofit = False
    tbl_draft.rows[0].cells[0].width = Inches(6.7)
    set_table_borders(tbl_draft, color="CCCCCC", sz="4")

    draft_parts = [
        "1. Hook (Chào hỏi & Câu hỏi mở đầu gây chú ý): .......................................................................................................................................................",
        "2. Location (Vị trí tọa lạc ở đâu tại TP.HCM): .........................................................................................................................................................",
        "3. History & Architecture (Năm xây dựng, nét đặc sắc kiến trúc): ......................................................................................................................................\n......................................................................................................................................................................................................................................",
        "4. Fun Fact (Một chi tiết thú vị, độc đáo về địa danh): ..................................................................................................................................................",
        "5. Advice & Closing (Lời khuyên cho du khách & Lời chào tạm biệt): .................................................................................................................................\n[ Tổng số từ của kịch bản nhóm: ........ từ (Yêu cầu: 80 - 110 từ) ]"
    ]

    for idx, d_txt in enumerate(draft_parts):
        cell = tbl_draft.rows[idx].cells[0]
        cell.width = Inches(6.7)
        set_cell_margins(cell, 25, 25, 35, 35)
        p = cell.paragraphs[0]; r = p.add_run(d_txt); r.font.size = Pt(8.5)
        if idx == 4: r.font.bold = True

def add_step3_to_5(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 4: NHIỆM VỤ 3 – LUYỆN PHÁT ÂM VÀ THU ÂM TIẾNG ANH (AUDIO GUIDE)"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl_tech = doc.add_table(rows=1, cols=1)
    tbl_tech.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_tech.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_tech, border_color="006600", bg_color="F2FAF2")
    p_t = tbl_tech.rows[0].cells[0].paragraphs[0]
    p_t.add_run(
        "• Tiêu chuẩn phát âm: Bật rõ các âm đuôi (/s/, /z/, /ed/), ngắt nghỉ đúng cụm nghĩa, ngữ điệu thân thiện, tự tin.\n"
        "• Quy chuẩn âm học khi thu âm: Thực hiện trong phòng kín, kéo rèm để chống vang (độ ồn < 45dB); cầm điện thoại cách miệng 15 cm chếch góc 45 độ (không thở trực tiếp vào micro làm rè tiếng).\n"
        "• Thời lượng chuẩn: 45 - 60 giây (tệp định dạng MP3 hoặc M4A phổ biến trên điện thoại)."
    )

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PHẦN 5: NHIỆM VỤ 4 & 5 – NỘP BÀI DUYỆT, THUYẾT TRÌNH VÀ TƯƠNG TÁC SỐ"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl_sub = doc.add_table(rows=4, cols=2)
    tbl_sub.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sub.autofit = False
    set_table_borders(tbl_sub, color="B0C4DE", sz="4")

    steps_end = [
        ("1. Nộp bài lên Cổng học liệu số", "Truy cập https://english8-two.vercel.app/ → Nhấn nút 'Student Submit / Học sinh nộp bài' → Tải lên ảnh, dán bài viết tiếng Anh và file audio của nhóm."),
        ("2. Giáo viên kiểm duyệt (Teacher Verification)", "Giáo viên thẩm định bài viết, nghe bản thu âm và bấm Phê duyệt công bố chính thức sản phẩm với huy hiệu 'Thông tin đã được phê duyệt ✓'."),
        ("3. Thuyết trình trước lớp (Tiết 2)", "Đại diện nhóm lên bục giảng thuyết trình; giáo viên bật Chế độ thuyết trình (Presentation Mode) trên máy chiếu để chiếu bài và phát file audio của nhóm; GV chấm điểm."),
        ("4. Quét mã QR & Tương tác số", "Cả lớp dùng điện thoại quét mã QR trên màn hình máy chiếu để nghe lại, bấm '❤️ Thả tim' và gửi bình luận nhận xét trực tiếp (hiển thị tức thì trên màn hình máy chiếu).")
    ]

    for idx, (s_t, s_d) in enumerate(steps_end):
        row = tbl_sub.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.3); c1.width = Inches(4.4)
        set_cell_margins(c0, 30, 30, 40, 40); set_cell_margins(c1, 30, 30, 40, 40)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(8.5)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(8.5)

def build_all_worksheets(filename=None):
    if filename is None:
        target_dir = os.path.join("01_Ho_So_Du_Thi_SGDDT", "Folder 2 - Ke hoach bai day")
        os.makedirs(target_dir, exist_ok=True)
        filename = os.path.join(target_dir, "Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx")

    doc = docx.Document()
    setup_document_styles(doc)

    groups = [
        {
            "num": 1,
            "site_vn": "Bưu điện Trung tâm Sài Gòn",
            "site_en": "Saigon Central Post Office",
            "main_hint": "Toàn cảnh mặt tiền màu vàng rực rỡ với chiếc đồng hồ lớn cổ kính ở giữa cổng chính nhìn từ Công xã Paris.",
            "sub_hints": [
                "Vòm mái trần uốn cong bằng khung sắt tuyệt đẹp theo phong cách kiến trúc Pháp cổ điển bên trong.",
                "Hai tấm bản đồ lịch sử vẽ tay trên tường gỗ ở hai bên lối vào chính.",
                "Các bốt điện thoại công cộng bằng gỗ cổ kính vẫn được gìn giữ nguyên vẹn.",
                "Quầy viết thư tay và gửi bưu thiếp truyền thống phục vụ du khách khắp nơi trên thế giới."
            ],
            "vocab": "classic architecture (/ˈklæsɪk ˈɑːkɪtektʃə/), arched ceiling (/ɑːtʃt ˈsiːlɪŋ/), telephone booths (/ˈtelɪfəʊn buːðz/), traditional postcards (/trəˈdɪʃənl ˈpəʊstkɑːdz/), European style (/ˌjʊərəˈpiːən staɪl/).",
            "grammar": "was built between 1886 and 1891; Inside, you can see...; Many visitors still come here to...; You should take photos and keep the place clean...",
            "template": "1. Hook: Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station?\n2. Location: It is the Saigon Central Post Office, located in the heart of District 1, right next to Notre Dame Cathedral.\n3. History/Architecture: This famous monument was built between 1886 and 1891 during the French period. Inside, you can see a beautiful arched ceiling and old wooden telephone booths.\n4. Fun Fact: Here is a fun fact: Many visitors still come here every day to write and send traditional postcards to their families around the world!\n5. Advice: You should take some nice photos here and remember to keep the place clean. Thank you and enjoy your trip in Saigon!"
        },
        {
            "num": 2,
            "site_vn": "Nhà thờ Đức Bà Sài Gòn",
            "site_en": "Notre Dame Cathedral of Saigon",
            "main_hint": "Toàn cảnh hai tháp chuông uy nghiêm cao gần 60 mét nổi bật trên nền trời xanh của trung tâm Quận 1.",
            "sub_hints": [
                "Bức tường gạch đỏ Marseille nguyên bản không cần trát vữa nhưng vẫn giữ nguyên màu sắc tươi sáng.",
                "Tượng Đức Mẹ Hòa Bình bằng đá cẩm thạch trắng đặt trang nghiêm tại công viên trước thánh đường.",
                "Đôi tháp chuông với đỉnh tháp nhọn gắn thánh giá cao vút.",
                "Những ô cửa kính màu tuyệt đẹp phản chiếu ánh sáng lung linh bên trong giáo đường."
            ],
            "vocab": "cathedral (/kəˈθiːdrəl/), bell towers (/bel ˈtaʊəz/), red bricks (/red brɪks/), peaceful square (/ˈpiːsfl skweə/), stained glass windows (/steɪnd ɡlɑːs ˈwɪndəʊz/).",
            "grammar": "was established in the late 19th century; All building materials were imported from France; You should visit respectfully...",
            "template": "1. Hook: Welcome to the most iconic religious landmark of Ho Chi Minh City: Notre Dame Cathedral!\n2. Location: Located right at Paris Square in District 1, it stands proudly opposite the Central Post Office.\n3. History/Architecture: Completed in 1880, all red bricks were imported directly from Marseille, France, and its two bell towers reach nearly 60 meters high.\n4. Fun Fact: Did you know? In front of the church stands the peaceful statue of Our Lady of Peace, bringing calm vibes to the busy city.\n5. Advice: Although it is currently being renovated, you should take wonderful photos outside and enjoy the friendly pigeon square. Have a great day!"
        },
        {
            "num": 3,
            "site_vn": "Chợ Bến Thành",
            "site_en": "Ben Thanh Market",
            "main_hint": "Toàn cảnh tháp đồng hồ 4 mặt ở Cửa Nam chợ Bến Thành – biểu tượng văn hóa và giao thương sôi động bậc nhất của TP.HCM.",
            "sub_hints": [
                "Bức phù điêu đắp nổi tinh xảo hình con bò tót hoặc con cá trên cổng chợ.",
                "Gian hàng bày bán đồ thủ công mỹ nghệ, nón lá, áo dài, túi thổ cẩm truyền thống.",
                "Khu ẩm thực náo nhiệt với các món đặc sản Nam Bộ (bánh xèo, chè ba màu, cà phê vợt).",
                "Tiểu thương đang trò chuyện, giao tiếp tiếng Anh vui vẻ và hướng dẫn du khách mua sắm thân thiện."
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
            "vocab": "historic wharf (/hɪˈstɒrɪk wɔːf/), ceramic dragons (/səˈræmɪk ˈdræɡənz/), patriot (/ˈpeɪtriət/), set sail (/set seɪl/), national independence (/ˈnæʃnəl ˌɪndɪˈpendəns/), precious artifacts (/ˈpreʃəs ˈɑːtɪfækts/).",
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

    try:
        doc.save(filename)
        print(f"File {filename} created successfully.")
    except PermissionError:
        alt_name = filename.replace(".docx", "_CapNhatWeb.docx")
        doc.save(alt_name)
        print(f"File {filename} is currently open in Word. Saved updated version as {alt_name} instead.")

if __name__ == "__main__":
    build_all_worksheets()
