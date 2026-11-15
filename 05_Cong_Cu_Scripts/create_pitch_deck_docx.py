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

def build_pitch_deck(filename="Ban_Demo_Y_Tuong_Saigon_Heritage_Guide.docx"):
    doc = docx.Document()
    setup_document_styles(doc)

    # ================= COVER PAGE =================
    p_pre = doc.add_paragraph()
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pre.paragraph_format.space_before = Pt(16)
    p_pre.paragraph_format.space_after = Pt(2)
    r = p_pre.add_run("BẢN ĐỀ XUẤT Ý TƯỞNG DỰ ÁN & DEMO HỆ THỐNG SỐ HÓA\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(6)
    r = p_title.add_run("SAIGON HERITAGE QR AUDIO GUIDE\n")
    r.font.bold = True; r.font.size = Pt(22); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49) # Navy
    r2 = p_title.add_run("VOICES OF OUR COMMUNITY\n")
    r2.font.bold = True; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(0xC8, 0x10, 0x2E) # Accent Red

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(16)
    r = p_sub.add_run("Giải pháp Chuyển đổi số Di sản TP.HCM & Dạy học STEM Tiếng Anh Tương tác\n(Interactive Digital Heritage & Youth Audio Guide Portal)")
    r.font.italic = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # Executive Overview Box
    tbl_sum = doc.add_table(rows=1, cols=1)
    tbl_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sum.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_sum, border_color="002B49", bg_color="F4F7FA")
    p_box = tbl_sum.rows[0].cells[0].paragraphs[0]
    r = p_box.add_run("TÓM TẮT DỰ ÁN DÀNH CHO ĐỐI TÁC / KHÁCH HÀNG:\n")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)
    r = p_box.add_run(
        "• Mục tiêu cốt lõi: Số hóa trải nghiệm du lịch di sản TP.HCM thông qua hệ sinh thái hướng dẫn âm thanh (Audio Guide) tiếng Anh 45 - 60 giây và mã QR thông minh được nghiên cứu, chụp ảnh, thu âm và xuất bản bởi chính các nhóm học sinh trung học.\n"
        "• Sản phẩm bàn giao: 01 Nền tảng Website Portal trực quan (kèm bộ công cụ tải ảnh 1 chính + 4 phụ và tự tạo QR code), 04 Trạm di tích điểm mẫu hoàn chỉnh, Bộ ấn phẩm Bảng thông tin A4/A5 dán tại di tích và Bộ phiếu quy trình học tập STEM.\n"
        "• Tính ứng dụng: Phục vụ trực tiếp cho du khách quốc tế tham quan tự túc, đồng thời làm kho học liệu số dùng chung cho ngành giáo dục và quảng bá văn hóa địa phương."
    )
    r.font.size = Pt(10)

    doc.add_page_break()

    # ================= SECTION I: CONTEXT & PROBLEM =================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8); h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("I. BỐI CẢNH & BÀI TOÁN THỰC TIỄN (PROBLEM & OPPORTUNITY)"); r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    p = doc.add_paragraph()
    p.add_run(
        "1. Thực trạng tại các di tích lịch sử - văn hóa TP.HCM:\n"
        "Thành phố Hồ Chí Minh sở hữu kho tàng di tích đồ sộ với hàng triệu lượt khách quốc tế mỗi năm. Tuy nhiên, bảng thông tin tại các điểm đến hiện nay thường gặp 3 hạn chế lớn: "
        "(1) Chỉ có tiếng Việt hoặc bản dịch tiếng Anh quá dài dòng, đơn điệu; "
        "(2) Thiếu công nghệ số tương tác âm thanh trực quan trên smartphone; "
        "(3) Chưa có sự tham gia của thế hệ trẻ trong việc lan tỏa câu chuyện văn hóa quê hương.\n\n"
        "2. Giải pháp đột phá từ Dự án Saigon Heritage Audio Guide:\n"
        "Biến mỗi di tích thành một 'Trạm âm thanh thông minh'. Du khách chỉ cần quét mã QR bằng camera điện thoại là có thể lắng nghe câu chuyện lịch sử - kiến trúc súc tích trong 45 - 60 giây bằng tiếng Anh với giọng đọc trong trẻo của các 'Đại sứ văn hóa trẻ', kèm hình ảnh chi tiết và bản phụ đề chữ lớn."
    )

    # ================= SECTION II: SCREENSHOTS SHOWCASE =================
    h_screen = doc.add_paragraph()
    h_screen.paragraph_format.space_before = Pt(12); h_screen.paragraph_format.space_after = Pt(4)
    r = h_screen.add_run("II. HÌNH ẢNH GIAO DIỆN NỀN TẢNG WEB DEMO THỰC TẾ (LIVE PLATFORM UI SHOWCASE)"); r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_after = Pt(8)
    r = p_note.add_run("Dưới đây là các hình ảnh chụp thực tế từ hệ thống Website Cổng học liệu số (demo_khach_xem.html) được thiết kế hiện đại, responsive và sẵn sàng triển khai:")
    r.font.italic = True; r.font.size = Pt(10)

    # Image 1: Hero Banner & Metrics
    p_img1 = doc.add_paragraph(); p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_img1.add_run().add_picture('screenshot_hero.png', width=Inches(6.5))
        p_cap1 = doc.add_paragraph(); p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap1.paragraph_format.space_after = Pt(12)
        r = p_cap1.add_run("Hình 1: Giao diện Trang chủ, Banner Dự án và Bộ 4 Chỉ số Trụ cột Học liệu số")
        r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    except Exception as e:
        p_img1.add_run(f"[Hình 1: Screenshot Hero Banner - {e}]")

    # Image 2: Cards Overview
    p_img2 = doc.add_paragraph(); p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_img2.add_run().add_picture('screenshot_cards_top.png', width=Inches(6.5))
        p_cap2 = doc.add_paragraph(); p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.paragraph_format.space_after = Pt(12)
        r = p_cap2.add_run("Hình 2: Cận cảnh Thẻ Di tích Đa phương tiện (Ảnh chính, Hàng 4 ảnh phụ, Audio Player & Kịch bản)")
        r.font.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    except Exception as e:
        p_img2.add_run(f"[Hình 2: Screenshot Cards - {e}]")

    doc.add_page_break()

    # Image 3 & 4: Modal Submit & QR Popup in 2 columns
    h_sub_title = doc.add_paragraph()
    h_sub_title.paragraph_format.space_before = Pt(6); h_sub_title.paragraph_format.space_after = Pt(4)
    r = h_sub_title.add_run("CÔNG CỤ NỘP BÀI CỦA HỌC SINH VÀ TRẢI NGHIỆM QUÉT MÃ QR DU KHÁCH"); r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    tbl_modals = doc.add_table(rows=1, cols=2)
    tbl_modals.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_modals.autofit = False
    c_m1, c_m2 = tbl_modals.rows[0].cells[0], tbl_modals.rows[0].cells[1]
    c_m1.width = Inches(4.0); c_m2.width = Inches(2.7)
    set_cell_margins(c_m1, 40, 40, 40, 40); set_cell_margins(c_m2, 40, 40, 40, 40)
    set_cell_background(c_m1, "FAFAFA"); set_cell_background(c_m2, "FAFAFA")

    p_m1 = c_m1.paragraphs[0]; p_m1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_m1.add_run().add_picture('screenshot_modal_submit_crop.png', width=Inches(3.7))
        r = c_m1.add_paragraph(); r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = r.add_run("Hình 3: Form học sinh tự tải 1 ảnh chính + 4 ảnh phụ & nhập kịch bản")
        r1.font.bold = True; r1.font.size = Pt(8.5); r1.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    except Exception:
        p_m1.add_run("[Hình 3: Form nộp bài]")

    p_m2 = c_m2.paragraphs[0]; p_m2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        p_m2.add_run().add_picture('screenshot_modal_qr_crop.png', width=Inches(2.4))
        r = c_m2.add_paragraph(); r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = r.add_run("Hình 4: Popup phóng to mã QR & Tải mã QR về in ấn")
        r2.font.bold = True; r2.font.size = Pt(8.5); r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    except Exception:
        p_m2.add_run("[Hình 4: Popup mã QR]")

    # ================= SECTION III: 4 LANDMARKS DEMO =================
    doc.add_page_break()
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8); h3.paragraph_format.space_after = Pt(4)
    r = h3.add_run("III. NỘI DUNG CHI TIẾT 4 DI TÍCH TRỌNG ĐIỂM (LANDMARK SHOWCASE)"); r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    landmarks = [
        {
            "num": 1,
            "title_vn": "BƯU ĐIỆN TRUNG TÂM SÀI GÒN",
            "title_en": "Saigon Central Post Office (Quận 1)",
            "img": "post_office.jpg",
            "qr": "qr_post_office.png",
            "script": "“Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station? It is the Saigon Central Post Office, located in the heart of District 1, right next to Notre Dame Cathedral. This famous monument was built between 1886 and 1891 during the French period. Inside, you can see a magnificent arched ceiling and historic wooden telephone booths. Many visitors still come here every day to write and send traditional postcards worldwide! You should take some nice photos here and keep the place clean. Enjoy your visit in Saigon!”",
            "vn": "Bưu điện Trung tâm Sài Gòn là công trình kiến trúc biểu tượng tuyệt đẹp xây dựng từ 1886 - 1891, nổi bật với mái vòm cong lớn kiểu châu Âu và những quầy điện thoại gỗ cổ kính vẫn còn được bảo tồn nguyên vẹn.",
            "photos_desc": "• Ảnh chính: Mặt tiền vàng kem cổ điển | • 4 ảnh phụ: Mái vòm sắt bánh mì, Bốt điện thoại gỗ, Bản đồ cổ viễn thông, Quầy viết thư tay."
        },
        {
            "num": 2,
            "title_vn": "NHÀ THỜ ĐỨC BÀ SÀI GÒN",
            "title_en": "Notre Dame Cathedral of Saigon (Quận 1)",
            "img": "notre_dame.jpg",
            "qr": "qr_notre_dame.png",
            "script": "“Standing gracefully in Paris Square, Notre Dame Cathedral is a famous architectural symbol of Ho Chi Minh City. It was constructed between 1877 and 1880 with all red bricks imported directly from Marseille, France. The cathedral has two impressive bell towers that reach nearly 60 meters into the sky. Fun fact: The exterior red bricks have kept their bright color for more than 140 years without any plastering! When visiting, you should speak quietly to preserve the peaceful atmosphere. Have a wonderful visit!”",
            "vn": "Nhà thờ Đức Bà được xây dựng từ năm 1877 với toàn bộ gạch đỏ vận chuyển từ Marseille (Pháp), nổi bật với 2 tháp chuông cao 60m không hề bám rêu mốc sau hơn 140 năm.",
            "photos_desc": "• Ảnh chính: Toàn cảnh 2 tháp chuông | • 4 ảnh phụ: Đỉnh tháp chuông vút cao, Cận cảnh gạch đỏ Marseille, Tượng Đức Mẹ cẩm thạch trắng, Phố sách bên hông."
        },
        {
            "num": 3,
            "title_vn": "CHỢ BẾN THÀNH",
            "title_en": "Ben Thanh Market (Quận 1)",
            "img": "ben_thanh.jpg",
            "qr": "qr_ben_thanh.png",
            "script": "“Welcome to Ben Thanh Market, the most vibrant trading landmark in Ho Chi Minh City! Opened in 1914, its iconic south clock tower has become a beloved symbol recognized worldwide. Inside, you can discover hundreds of lively stalls selling local food, silk, coffee, and traditional handicrafts. Here is a helpful tip: Don't miss trying a bowl of authentic southern noodle soup or fresh tropical fruit desserts! You should also practice friendly bargaining with the local vendors. Enjoy the colorful sights, sounds, and flavors of Saigon!”",
            "vn": "Chợ Bến Thành khánh thành năm 1914, nổi tiếng với tháp đồng hồ bốn mặt ở Cửa Nam. Nơi đây hội tụ hàng trăm gian hàng thủ công truyền thống và thiên đường ẩm thực Nam Bộ.",
            "photos_desc": "• Ảnh chính: Tháp đồng hồ Cửa Nam | • 4 ảnh phụ: Bức phù điêu đắp nổi, Gian hàng nón lá thủ công, Khu ẩm thực chè/bún, Khách quốc tế mua sắm."
        },
        {
            "num": 4,
            "title_vn": "BẾN NHÀ RỒNG - BẢO TÀNG HỒ CHÍ MINH",
            "title_en": "Dragon Wharf - Ho Chi Minh Museum (Quận 4)",
            "img": "dragon_wharf.jpg",
            "qr": "qr_dragon_wharf.png",
            "script": "“Welcome to Dragon Wharf, officially known as the Ho Chi Minh Museum, situated along the Saigon River in District 4. Originally built in 1863 as a commercial port headquarters, the rooftop features two distinct ceramic dragon statues. This site holds tremendous historical significance: On June 5th, 1911, young patriot Nguyen Tat Thanh set sail abroad to find a path for national liberation. Today, it preserves thousands of precious artifacts and photographs. Please visit respectfully and explore our history!”",
            "vn": "Bến Nhà Rồng xây dựng năm 1863 bên sông Sài Gòn, nổi bật với đôi rồng gốm trên nóc. Nơi ghi dấu sự kiện lịch sử Bác Hồ ra đi tìm đường cứu nước ngày 05/06/1911.",
            "photos_desc": "• Ảnh chính: Tòa nhà cổ nhìn từ sông | • 4 ảnh phụ: Đôi rồng gốm men xanh nóc nhà, Bia tưởng niệm 5/6/1911, Phòng hiện vật lịch sử, Cảng sông lộng gió."
        }
    ]

    for lm in landmarks:
        tbl_item = doc.add_table(rows=1, cols=2)
        tbl_item.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_item.autofit = False
        c_left, c_right = tbl_item.rows[0].cells[0], tbl_item.rows[0].cells[1]
        c_left.width = Inches(3.0); c_right.width = Inches(3.7)
        set_cell_margins(c_left, 45, 45, 45, 45); set_cell_margins(c_right, 45, 45, 45, 45)
        set_cell_background(c_left, "F9FAFC"); set_cell_background(c_right, "FFFFFF")

        # Left Column: Landmark Image + QR code
        p_img = c_left.paragraphs[0]; p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            p_img.add_run().add_picture(lm["img"], width=Inches(2.8))
        except Exception:
            p_img.add_run("[Ảnh di tích]")

        p_qr = c_left.add_paragraph(); p_qr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_qr.paragraph_format.space_before = Pt(3)
        try:
            p_qr.add_run().add_picture(lm["qr"], width=Inches(1.15))
            r = p_qr.add_run("\nQuét nghe Audio Guide thử nghiệm")
            r.font.size = Pt(8); r.font.italic = True; r.font.color.rgb = RGBColor(0xC8, 0x10, 0x2E)
        except Exception:
            p_qr.add_run("[Mã QR]")

        # Right Column: Content & Script
        p_txt = c_right.paragraphs[0]
        r = p_txt.add_run(f"TRẠM {lm['num']}: {lm['title_vn']}\n"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)
        r = p_txt.add_run(f"{lm['title_en']}\n\n"); r.font.italic = True; r.font.size = Pt(9)

        r = p_txt.add_run("Kịch bản Tiếng Anh (English Script - 5 Steps):\n"); r.font.bold = True; r.font.size = Pt(9)
        r = p_txt.add_run(f"{lm['script']}\n\n"); r.font.size = Pt(8.5); r.font.italic = True

        r = p_txt.add_run("Tóm tắt tiếng Việt (Chữ lớn trợ năng):\n"); r.font.bold = True; r.font.size = Pt(8.5)
        r = p_txt.add_run(f"{lm['vn']}\n\n"); r.font.size = Pt(8)

        r = p_txt.add_run("Quy cách bộ ảnh 1+4:\n"); r.font.bold = True; r.font.size = Pt(8.5)
        r = p_txt.add_run(f"{lm['photos_desc']}"); r.font.size = Pt(8); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(2); p_spacer.paragraph_format.space_after = Pt(4)

    doc.add_page_break()

    # ================= SECTION IV: VALUE PROPOSITION =================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8); h4.paragraph_format.space_after = Pt(4)
    r = h4.add_run("IV. GIÁ TRỊ MANG LẠI CHO ĐỐI TÁC & KHÁCH HÀNG (VALUE PROPOSITION)"); r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    tbl_val = doc.add_table(rows=5, cols=3)
    tbl_val.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_val.autofit = False
    set_table_borders(tbl_val, color="B0C4DE", sz="4")

    val_headers = ["Đối tượng thụ hưởng", "Giá trị mang lại", "Chỉ số đo lường (KPIs)"]
    val_data = [
        ("Đối với Nhà trường & Giáo viên", "• Có ngay một giải pháp dạy học số đạt chuẩn Kế hoạch 12070 của Sở GD&ĐT TP.HCM.\n• Đổi mới phương pháp dạy học STEM môn Tiếng Anh thực chất, tạo ấn tượng với Ban giám khảo và phụ huynh.", "100% học sinh hào hứng tham gia; hồ sơ dự thi đủ điều kiện tranh giải cấp Thành phố."),
        ("Đối với Học sinh", "• Xóa bỏ rào cản sợ nói tiếng Anh thông qua việc đóng vai 'Đại sứ văn hóa nhí'.\n• Làm chủ kỹ năng số: xử lý âm thanh, tạo mã QR, nhiếp ảnh, thiết kế đồ họa Canva.", "Tỷ lệ nói trôi chảy đạt ≥ 85%; thành thạo 4 công cụ số phổ biến hiện nay."),
        ("Đối với Du khách Quốc tế", "• Tiếp cận câu chuyện văn hóa Sài Gòn một cách tức thì, tiện lợi chỉ với 1 thao tác quét mã trên smartphone, không cần cài ứng dụng phức tạp.", "Thời gian tiếp cận thông tin ≤ 3 giây; thời lượng nghe lý tưởng 45 - 60 giây."),
        ("Đối với Ngành Du lịch & Cộng đồng", "• Tạo ra kho học liệu số cộng đồng mở, có thể đóng góp vào Cổng thông tin du lịch thông minh của TP.HCM và ngành GD&ĐT.", "Bộ sản phẩm lưu trữ vĩnh viễn trên đám mây, dễ dàng mở rộng thêm các di tích mới.")
    ]

    for c_i, h in enumerate(val_headers):
        cell = tbl_val.rows[0].cells[c_i]
        set_cell_background(cell, "002B49")
        set_cell_margins(cell, 45, 45, 45, 45)
        p = cell.paragraphs[0]; r = p.add_run(h); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    tbl_val.rows[0].cells[0].width = Inches(1.8)
    tbl_val.rows[0].cells[1].width = Inches(3.2)
    tbl_val.rows[0].cells[2].width = Inches(1.7)

    for r_i, r_data in enumerate(val_data):
        row = tbl_val.rows[r_i + 1]
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            cell.width = [Inches(1.8), Inches(3.2), Inches(1.7)][c_i]
            set_cell_margins(cell, 40, 40, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(8.5)
            if c_i == 0: r.font.bold = True

    # Section V: Roadmap & Signature
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(12); h5.paragraph_format.space_after = Pt(4)
    r = h5.add_run("V. LỘ TRÌNH TRIỂN KHAI & ĐỀ XUẤT HỢP TÁC (PROJECT ROADMAP)"); r.font.bold = True; r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x00, 0x2B, 0x49)

    p_road = doc.add_paragraph()
    p_road.add_run(
        "• Giai đoạn 1 (Tuần 1): Khảo sát thực tế 4 di tích, chụp bộ 5 ảnh và soạn thảo kịch bản tiếng Anh 5 bước theo khung gợi ý.\n"
        "• Giai đoạn 2 (Tuần 2): Thu âm giọng đọc chuẩn âm học (S), xử lý lọc ồn và xuất bản lên Website Cổng di sản Sài Gòn (index.html).\n"
        "• Giai đoạn 3 (Tuần 3): In ấn bảng mica/formex khổ A4 kèm mã QR; tổ chức Triển lãm thực nghiệm Gallery Walk tại lớp học hoặc tại khuôn viên trường.\n"
        "• Giai đoạn 4 (Tuần 4): Hoàn thiện hồ sơ minh chứng dự thi và chuyển giao đường link Cổng học liệu số cho các đơn vị đối tác / Sở GD&ĐT TP.HCM."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sign.autofit = False
    tbl_sign.rows[0].cells[0].width = Inches(3.2); tbl_sign.rows[0].cells[1].width = Inches(3.5)

    c0, c1 = tbl_sign.rows[0].cells[0], tbl_sign.rows[0].cells[1]
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ĐẠI DIỆN ĐỐI TÁC / KHÁCH HÀNG\n"); r.font.bold = True
    r = p.add_run("(Ký và ghi rõ họ tên)"); r.font.italic = True

    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TP. Hồ Chí Minh, ngày       tháng 11 năm 2026\n"); r.font.italic = True
    r = p.add_run("ĐƠN VỊ PHÁT TRIỂN Ý TƯỞNG DỰ ÁN\n"); r.font.bold = True
    r = p.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"); r.font.italic = True
    r = p.add_run("[Họ và tên tác giả / Trưởng nhóm dự án]"); r.font.bold = True

    try:
        doc.save(filename)
        print(f"File {filename} created successfully with embedded web screenshots!")
    except PermissionError:
        alt_name = "Ban_Demo_Y_Tuong_Saigon_Heritage_Guide_v2.docx"
        doc.save(alt_name)
        print(f"File {filename} is open. Saved as {alt_name} instead.")

if __name__ == "__main__":
    build_pitch_deck()
