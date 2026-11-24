import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6] # blank layout

    # Colors
    NAVY = RGBColor(0, 43, 73)
    GOLD = RGBColor(212, 175, 55)
    LIGHT_BG = RGBColor(248, 249, 250)
    WHITE = RGBColor(255, 255, 255)
    DARK = RGBColor(30, 41, 59)
    GRAY = RGBColor(100, 116, 139)
    EMERALD = RGBColor(16, 149, 106)
    AMBER = RGBColor(217, 119, 6)
    BORDER_GRAY = RGBColor(226, 232, 240)

    LOGO_PATH = 'images/logo_hoa_hiep.png'

    # Helper: Set background color of slide
    def set_slide_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    # Helper: Add header banner to content slides
    def add_slide_header(slide, title, category="HỘI THI THIẾT KẾ CHỦ ĐỀ STEM & DẠY HỌC SỐ CẤP THCS"):
        # Header background
        header_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.15))
        header_bg.fill.solid()
        header_bg.fill.fore_color.rgb = NAVY
        header_bg.line.fill.background()

        # Logo
        if os.path.exists(LOGO_PATH):
            slide.shapes.add_picture(LOGO_PATH, Inches(0.4), Inches(0.12), height=Inches(0.9))

        # Text in header
        tx_box = slide.shapes.add_textbox(Inches(1.45), Inches(0.12), Inches(11.4), Inches(0.9))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        p0.text = category.upper()
        p0.font.size = Pt(10)
        p0.font.bold = True
        p0.font.color.rgb = GOLD
        p0.font.name = "Arial"

        p1 = tf.add_paragraph()
        p1.text = title
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p1.font.name = "Arial"

        # Footer band
        footer_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(7.1), prs.slide_width, Inches(0.4))
        footer_bg.fill.solid()
        footer_bg.fill.fore_color.rgb = RGBColor(15, 23, 42)
        footer_bg.line.fill.background()

        ft_box = slide.shapes.add_textbox(Inches(0.5), Inches(7.12), Inches(12.33), Inches(0.35))
        ft_tf = ft_box.text_frame
        ft_tf.margin_left = ft_tf.margin_top = ft_tf.margin_right = ft_tf.margin_bottom = 0
        p_ft = ft_tf.paragraphs[0]
        p_ft.text = "TRƯỜNG THCS HÒA HIỆP • LỚP 8A2 • GVHD: CÔ TRẦN THỊ THANH NGUYÊN • https://english8-two.vercel.app/"
        p_ft.font.size = Pt(9.5)
        p_ft.font.color.rgb = RGBColor(148, 163, 184)
        p_ft.font.name = "Arial"

    # =========================================================================
    # SLIDE 1: TRANG BÌA GIỚI THIỆU THÔNG TIN ĐẦY ĐỦ (COVER SLIDE)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s1, NAVY)

    # Decorative top bar
    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.18))
    bar.fill.solid()
    bar.fill.fore_color.rgb = GOLD
    bar.line.fill.background()

    # Logo trường
    if os.path.exists(LOGO_PATH):
        s1.shapes.add_picture(LOGO_PATH, Inches(1.2), Inches(0.8), height=Inches(1.8))

    # Top Authority Info Box
    top_box = s1.shapes.add_textbox(Inches(3.3), Inches(0.8), Inches(8.8), Inches(1.8))
    tf1 = top_box.text_frame
    tf1.word_wrap = True

    p_dep = tf1.paragraphs[0]
    p_dep.text = "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH"
    p_dep.font.size = Pt(13)
    p_dep.font.bold = True
    p_dep.font.color.rgb = RGBColor(226, 232, 240)
    p_dep.font.name = "Arial"

    p_sch = tf1.add_paragraph()
    p_sch.text = "TRƯỜNG TRUNG HỌC CƠ SỞ HÒA HIỆP"
    p_sch.font.size = Pt(14)
    p_sch.font.bold = True
    p_sch.font.color.rgb = GOLD
    p_sch.font.name = "Arial"

    p_contest = tf1.add_paragraph()
    p_contest.text = "HỘI THI THIẾT KẾ CHỦ ĐỀ STEM & DẠY HỌC SỐ CẤP THCS NĂM HỌC 2024 - 2025"
    p_contest.font.size = Pt(12)
    p_contest.font.bold = True
    p_contest.font.color.rgb = RGBColor(148, 163, 184)
    p_contest.font.name = "Arial"

    # Main Project Title Box
    main_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(2.9), Inches(10.93), Inches(2.2))
    main_box.fill.solid()
    main_box.fill.fore_color.rgb = RGBColor(10, 56, 92)
    main_box.line.color.rgb = GOLD
    main_box.line.width = Pt(2)

    m_tf = main_box.text_frame
    m_tf.word_wrap = True
    m_tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p_sub = m_tf.paragraphs[0]
    p_sub.text = "BÁO CÁO KẾT QUẢ TRIỂN KHAI CHỦ ĐỀ STEM & DẠY HỌC SỐ MÔN TIẾNG ANH 8"
    p_sub.font.size = Pt(12)
    p_sub.font.bold = True
    p_sub.font.color.rgb = GOLD
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.name = "Arial"

    p_title = m_tf.add_paragraph()
    p_title.text = "SAIGON HERITAGE DIGITAL QR AUDIO GUIDE"
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE
    p_title.alignment = PP_ALIGN.CENTER
    p_title.font.name = "Arial"

    p_vn = m_tf.add_paragraph()
    p_vn.text = "HƯỚNG DẪN VIÊN SỐ GIỚI THIỆU DI SẢN LỊCH SỬ SÀI GÒN BẰNG TIẾNG ANH"
    p_vn.font.size = Pt(14)
    p_vn.font.bold = True
    p_vn.font.color.rgb = RGBColor(241, 245, 249)
    p_vn.alignment = PP_ALIGN.CENTER
    p_vn.font.name = "Arial"

    # 3 Metadata Badges at Bottom
    card_w = Inches(3.45)
    gap = Inches(0.29)
    y_cards = Inches(5.35)

    # Card 1: Giáo viên & Lớp
    c1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), y_cards, card_w, Inches(1.5))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(15, 34, 56)
    c1.line.color.rgb = RGBColor(51, 65, 85)
    tf_c1 = c1.text_frame
    tf_c1.word_wrap = True
    p = tf_c1.paragraphs[0]
    p.text = "👤 THÔNG TIN THỰC HIỆN"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p = tf_c1.add_paragraph()
    p.text = "• GVHD: Cô Trần Thị Thanh Nguyên\n• Lớp thực hiện: Lớp 8A2 (4 Nhóm)\n• Môn học: Tiếng Anh 8 (Global Success)"
    p.font.size = Pt(11)
    p.font.color.rgb = WHITE

    # Card 2: Quy chuẩn Sư phạm
    c2 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2) + card_w + gap, y_cards, card_w, Inches(1.5))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(15, 34, 56)
    c2.line.color.rgb = RGBColor(51, 65, 85)
    tf_c2 = c2.text_frame
    tf_c2.word_wrap = True
    p = tf_c2.paragraphs[0]
    p.text = "📚 CẤU TRÚC SƯ PHẠM CHUẨN"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p = tf_c2.add_paragraph()
    p.text = "• Công văn 5512/BGDĐT-GDTrH (4 Hoạt động)\n• Kế hoạch 12070/KH-SGDĐT TP.HCM\n• Khung STEM & Chuyển đổi số giáo dục"
    p.font.size = Pt(11)
    p.font.color.rgb = WHITE

    # Card 3: Sản phẩm Số
    c3 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2) + (card_w + gap)*2, y_cards, card_w, Inches(1.5))
    c3.fill.solid()
    c3.fill.fore_color.rgb = RGBColor(15, 34, 56)
    c3.line.color.rgb = RGBColor(51, 65, 85)
    tf_c3 = c3.text_frame
    tf_c3.word_wrap = True
    p = tf_c3.paragraphs[0]
    p.text = "🌐 NỀN TẢNG SẢN PHẨM SỐ"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p = tf_c3.add_paragraph()
    p.text = "• Website: https://english8-two.vercel.app/\n• Hệ thống 4 Trạm Audio Guide & QR Code\n• Tương tác: Trắc nghiệm, Thả tim, Bình luận"
    p.font.size = Pt(11)
    p.font.color.rgb = WHITE

    # =========================================================================
    # SLIDE 2: MỤC TIÊU DẠY HỌC & TÍCH HỢP STEM
    # =========================================================================
    s2 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s2, LIGHT_BG)
    add_slide_header(s2, "MỤC TIÊU DẠY HỌC & Ý NGHĨA TÍCH HỢP GIÁO DỤC STEM")

    # 4 STEM Pillars Cards
    stem_w = Inches(2.8)
    stem_gap = Inches(0.24)
    stem_y = Inches(1.4)
    stem_h = Inches(5.4)

    pillars = [
        ("SCIENCE (Khoa học)", "Lịch sử & Văn hóa", EMERALD, [
            "Tìm hiểu lịch sử, niên đại và ý nghĩa văn hóa của 4 di tích biểu tượng TP.HCM.",
            "Nghiên cứu nguyên lý vi khí hậu nhiệt đới, kiến trúc Gothic, Romanesque và vật liệu bền vững (gạch đỏ Marseille).",
            "Bảo tồn và lan tỏa niềm tự hào di sản đến bạn bè quốc tế."
        ]),
        ("TECHNOLOGY (Công nghệ)", "Âm thanh & Số hóa", NAVY, [
            "Học sinh sử dụng thiết bị số để ghi âm tiếng Anh, lọc tạp âm và chuẩn hóa file MP3.",
            "Xây dựng Web Portal đa phương tiện, tích hợp nghe trực tuyến và mã QR Code động.",
            "Hệ thống số hóa hỗ trợ học tập trực tiếp tại lớp và trải nghiệm từ xa."
        ]),
        ("ENGINEERING (Kỹ thuật)", "Thiết kế Poster & Quy trình", AMBER, [
            "Quy trình 5 vai trò làm việc nhóm chuẩn STEM rõ ràng, trách nhiệm minh bạch.",
            "Thiết kế bố cục Poster Infographic trực quan (1 ảnh bao quát + 4 góc chụp chi tiết).",
            "Kỹ thuật xây dựng kịch bản chuẩn 5 bước (Hook, Intro, History, Tip, Outro)."
        ]),
        ("MATHEMATICS (Toán học)", "Đo lường & Chuẩn hóa", RGBColor(124, 58, 237), [
            "Khống chế dung lượng bài thuyết minh chuẩn: 80 - 110 từ (Grade 8 Standard).",
            "Kiểm soát thời lượng đoạn ghi âm: 45 - 60 giây (chuẩn Audio Guide quốc tế).",
            "Thống kê số liệu niên đại (1863, 1877, 1886, 1914) và đánh giá trắc nghiệm 3 câu."
        ])
    ]

    for idx, (title, sub, col, items) in enumerate(pillars):
        x = Inches(0.6) + idx * (stem_w + stem_gap)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, stem_y, stem_w, stem_h)
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = col
        card.line.width = Pt(2)

        # Header tag on top of card
        tag = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.15), stem_y + Inches(0.15), stem_w - Inches(0.3), Inches(0.8))
        tag.fill.solid()
        tag.fill.fore_color.rgb = col
        tag.line.fill.background()
        tf = tag.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(9.5)
        p.font.color.rgb = RGBColor(241, 245, 249)
        p.alignment = PP_ALIGN.CENTER

        # Body items
        body = s2.shapes.add_textbox(x + Inches(0.15), stem_y + Inches(1.1), stem_w - Inches(0.3), Inches(4.1))
        b_tf = body.text_frame
        b_tf.word_wrap = True
        for i_idx, item in enumerate(items):
            p = b_tf.paragraphs[0] if i_idx == 0 else b_tf.add_paragraph()
            p.text = f"✔ {item}\n"
            p.font.size = Pt(10.5)
            p.font.color.rgb = DARK
            p.font.name = "Arial"

    # =========================================================================
    # SLIDE 3: TIẾN TRÌNH DẠY HỌC 4 HOẠT ĐỘNG (CHUẨN CV 5512)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s3, LIGHT_BG)
    add_slide_header(s3, "TIẾN TRÌNH DẠY HỌC 4 HOẠT ĐỘNG THEO CÔNG VĂN 5512/BGDĐT")

    act_w = Inches(5.9)
    act_h = Inches(2.6)
    gap_x = Inches(0.33)
    gap_y = Inches(0.25)

    activities = [
        ("HOẠT ĐỘNG 1: XÁC ĐỊNH VẤN ĐỀ & GIAO NHIỆM VỤ (Tiết 1)", [
            "• Giáo viên đặt vấn đề thực tiễn: Khách du lịch quốc tế đến TP.HCM cần thuyết minh tiếng Anh ngắn gọn, chính xác.",
            "• Giao nhiệm vụ cho 4 Nhóm lớp 8A2 thiết kế Hướng dẫn viên số (Audio Guide + Infographic QR Code) cho 4 địa danh.",
            "• Phân công 5 vai trò STEM trong mỗi nhóm: Trưởng nhóm, Tìm ảnh, Viết kịch bản, Thu âm tiếng Anh, Phụ trách số hóa."
        ], NAVY),
        ("HOẠT ĐỘNG 2: NGHIÊN CỨU KIẾN THỨC NỀN & THỰC HIỆN DỰ ÁN", [
            "• Học sinh tự tìm hiểu thông tin lịch sử, chọn 01 ảnh toàn cảnh và 03 - 04 ảnh góc chụp chi tiết đạt chuẩn.",
            "• Soạn thảo bài giới thiệu tiếng Anh 80 - 110 từ theo kịch bản 5 bước (sử dụng câu bị động, động từ khuyết thiếu).",
            "• Luyện phát âm ngữ điệu, thực hiện thu âm file âm thanh tiếng Anh chuẩn độ dài 45 - 60 giây ngoài giờ học."
        ], EMERALD),
        ("HOẠT ĐỘNG 3: BÁO CÁO THUYẾT TRÌNH TẠI LỚP (Tiết 2)", [
            "• Đại diện 4 nhóm trình bày sản phẩm bằng tiếng Anh tại lớp, kết hợp trình chiếu Poster và bật đoạn Audio Guide.",
            "• Thể hiện năng lực nói tiếng Anh lưu loát, tự tin trả lời các câu hỏi nghe hiểu và tương tác với các bạn.",
            "• Trải nghiệm học tập trực quan sinh động thông qua giao diện thuyết trình số hóa (Presentation Mode)."
        ], AMBER),
        ("HOẠT ĐỘNG 4: GIÁO VIÊN KIỂM DUYỆT & ĐÁNH GIÁ TƯƠNG TÁC SỐ", [
            "• Giáo viên mở bảng kiểm duyệt 'Teacher Verification' trên web để đánh giá dung lượng từ, ảnh thực tế và âm thanh.",
            "• GV trực tiếp tích chọn phê duyệt '✓ Đã Phê Duyệt' và nhập nhận xét đánh giá số hóa minh bạch.",
            "• Học sinh cả lớp dùng điện thoại/máy tính bảng quét mã QR trên màn hình, làm trắc nghiệm 3 câu, thả tim và bình luận chéo."
        ], RGBColor(124, 58, 237))
    ]

    for idx, (title, items, col) in enumerate(activities):
        row = idx // 2
        col_idx = idx % 2
        x = Inches(0.6) + col_idx * (act_w + gap_x)
        y = Inches(1.4) + row * (act_h + gap_y)

        box = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, act_w, act_h)
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = col
        box.line.width = Pt(2)

        # Title strip
        strip = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.12), y + Inches(0.12), act_w - Inches(0.24), Inches(0.42))
        strip.fill.solid()
        strip.fill.fore_color.rgb = col
        strip.line.fill.background()
        tf = strip.text_frame
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

        # Content
        body = s3.shapes.add_textbox(x + Inches(0.2), y + Inches(0.58), act_w - Inches(0.4), Inches(1.9))
        tf_b = body.text_frame
        tf_b.word_wrap = True
        for i_idx, itm in enumerate(items):
            p = tf_b.paragraphs[0] if i_idx == 0 else tf_b.add_paragraph()
            p.text = itm
            p.font.size = Pt(10.5)
            p.font.color.rgb = DARK
            p.font.name = "Arial"

    # =========================================================================
    # HELPER FOR GROUP SLIDES (SLIDES 4, 5, 6, 7)
    # =========================================================================
    def build_group_slide(prs, group_num, group_name, site_vn, site_en, duration_str, words_num, script_en, script_vn, cover_img, grammar_str):
        slide = prs.slides.add_slide(blank_slide_layout)
        set_slide_background(slide, LIGHT_BG)
        add_slide_header(slide, f"NHÓM {group_num}: {group_name.upper()} • {site_vn.upper()}")

        # Left Column: Product Info & Scripts (Width: 6.8 inches)
        left_w = Inches(6.8)
        left_x = Inches(0.6)

        # Info Box
        info_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, Inches(1.35), left_w, Inches(1.15))
        info_box.fill.solid()
        info_box.fill.fore_color.rgb = WHITE
        info_box.line.color.rgb = NAVY
        info_box.line.width = Pt(1.5)
        tf_info = info_box.text_frame
        tf_info.word_wrap = True

        p = tf_info.paragraphs[0]
        p.text = f"📍 {site_vn} ({site_en})"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = NAVY

        p = tf_info.add_paragraph()
        p.text = f"• Phụ trách: Nhóm {group_num} - {group_name} (Lớp 8A2)\n• Đoạn thu âm học sinh: {duration_str} • Dung lượng: {words_num} từ (Chuẩn Grade 8)\n• Ngữ pháp: {grammar_str}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = DARK

        # Script English Box
        scr_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, Inches(2.6), left_w, Inches(2.35))
        scr_box.fill.solid()
        scr_box.fill.fore_color.rgb = WHITE
        scr_box.line.color.rgb = EMERALD
        scr_box.line.width = Pt(1.5)
        tf_scr = scr_box.text_frame
        tf_scr.word_wrap = True

        p = tf_scr.paragraphs[0]
        p.text = f"🎙 KỊCH BẢN THUYẾT MINH TIẾNG ANH (AUDIO TRANSCRIPT - {words_num} WORDS):"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = EMERALD

        p = tf_scr.add_paragraph()
        p.text = f'"{script_en}"'
        p.font.size = Pt(11)
        p.font.italic = True
        p.font.color.rgb = NAVY

        # Script Vietnamese Box
        vn_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, Inches(5.05), left_w, Inches(1.85))
        vn_box.fill.solid()
        vn_box.fill.fore_color.rgb = RGBColor(241, 245, 249)
        vn_box.line.color.rgb = RGBColor(203, 213, 225)
        tf_vn = vn_box.text_frame
        tf_vn.word_wrap = True

        p = tf_vn.paragraphs[0]
        p.text = "📖 BẢN DỊCH TIẾNG VIỆT & NỘI DUNG TÌM HIỂU:"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = DARK

        p = tf_vn.add_paragraph()
        p.text = script_vn
        p.font.size = Pt(10.5)
        p.font.color.rgb = DARK

        # Right Column: 2 Visual Cards (1 Landmark photo + 1 Student Action Photo Placeholder)
        right_w = Inches(5.0)
        right_x = Inches(7.7)

        # Card 1: Landmark Photo
        if os.path.exists(cover_img):
            slide.shapes.add_picture(cover_img, right_x, Inches(1.35), width=right_w, height=Inches(2.55))
            # Caption badge
            cap = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, Inches(3.55), right_w, Inches(0.35))
            cap.fill.solid()
            cap.fill.fore_color.rgb = NAVY
            cap.line.fill.background()
            p = cap.text_frame.paragraphs[0]
            p.text = f"Ảnh tư liệu học sinh chọn lọc: {site_vn}"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

        # Card 2: Student Activity Placeholder Box (Vị trí chèn ảnh hoạt động nhóm)
        holder = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, Inches(4.05), right_w, Inches(2.85))
        holder.fill.solid()
        holder.fill.fore_color.rgb = RGBColor(254, 243, 199) # soft amber
        holder.line.color.rgb = AMBER
        holder.line.width = Pt(2)

        tf_h = holder.text_frame
        tf_h.word_wrap = True
        tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE

        p = tf_h.paragraphs[0]
        p.text = f"📷 VỊ TRÍ CHÈN ẢNH HOẠT ĐỘNG THỰC TẾ NHÓM {group_num}"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(180, 83, 9)
        p.alignment = PP_ALIGN.CENTER

        p = tf_h.add_paragraph()
        p.text = "(Ảnh học sinh thảo luận nhóm, ghi âm tiếng Anh trên điện thoại, hoặc thuyết trình tại lớp)\n\n👉 Thầy/Cô gửi ảnh cho em, em sẽ đưa ngay vào vị trí này!"
        p.font.size = Pt(10.5)
        p.font.color.rgb = DARK
        p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 4: NHÓM 1 - HERITAGE AMBASSADORS
    # =========================================================================
    build_group_slide(
        prs,
        group_num=1,
        group_name="Heritage Ambassadors",
        site_vn="Bưu điện Trung tâm Sài Gòn",
        site_en="Saigon Central Post Office",
        duration_str="49 giây",
        words_num="107",
        script_en="Hello and welcome to Ho Chi Minh City. Do you know which building looks like a classic European train station? It's Saigon Central Post Office, located in the heart of District 1, right next to Notre Dame Cathedral. This famous monument was built between 1886 and 1891 during the French period. Inside, you can see a magnificent arched ceiling and historic wooden telephone booths. Many visitors still come here everyday to write and send traditional postcards to their families around the world. You should take some nice photos here, and remember to keep the place clean. Thank you and enjoy your trip in Saigon!",
        script_vn="Bưu điện Trung tâm Sài Gòn là công trình kiến trúc biểu tượng tuyệt đẹp được xây dựng từ năm 1886 đến 1891 thời kỳ Pháp thuộc. Bên trong có vòm trần uốn cong tráng lệ và các buồng điện thoại cổ bằng gỗ. Du khách vẫn duy trì nét đẹp gửi bưu thiếp tay truyền thống cho người thân.",
        cover_img="images/post_office_main.jpg",
        grammar_str="Passive voice (was built between 1886 and 1891), Modal verb (You should take photos)"
    )

    # =========================================================================
    # SLIDE 5: NHÓM 2 - GOLDEN BELL
    # =========================================================================
    build_group_slide(
        prs,
        group_num=2,
        group_name="Golden Bell",
        site_vn="Nhà thờ Đức Bà Sài Gòn",
        site_en="Notre Dame Cathedral of Saigon",
        duration_str="48 giây",
        words_num="83",
        script_en="Standing gracefully in Paris Square, Notre Dame Cathedral is a famous architectural symbol of Ho Chi Minh City. It was constructed between 1877 and 1880 with all red bricks imported directly from Marseille, France. The cathedral has two impressive bell towers that reach nearly 60 meters into the sky. Fun fact: The exterior red bricks have kept their bright color for more than 140 years without any plastering. When visiting, you should speak quietly to preserve the peaceful atmosphere. Have a wonderful visit!",
        script_vn="Nhà thờ Đức Bà được xây dựng từ năm 1877 đến 1880 với toàn bộ gạch đỏ vận chuyển trực tiếp từ Marseille (Pháp), giữ màu sắc tươi sáng hơn 140 năm mà không cần tô trát. Hai tháp chuông cao gần 60m là biểu tượng tôn nghiêm giữa trung tâm thành phố.",
        cover_img="images/notre_dame_main.jpg",
        grammar_str="Passive voice (was constructed between 1877 and 1880), Modal verb (you should speak quietly)"
    )

    # =========================================================================
    # SLIDE 6: NHÓM 3 - SAIGON EXPLORERS
    # =========================================================================
    build_group_slide(
        prs,
        group_num=3,
        group_name="Saigon Explorers",
        site_vn="Chợ Bến Thành",
        site_en="Ben Thanh Market",
        duration_str="38 giây",
        words_num="81",
        script_en="Welcome to Ben Thanh Market, the most vibrant trading landmark in Ho Chi Minh City. Opened in 1914, its iconic south clock tower has become a beloved symbol recognized worldwide. Inside, you can discover hundreds of lively stalls selling local food, silk, coffee, and traditional handicrafts. Here is a helpful tip: Don't miss trying a bowl of authentic southern noodle soup or fresh tropical fruit dessert. You should also practice friendly bargaining with local vendors. Enjoy the colorful sights, sounds, and flavors of Saigon!",
        script_vn="Chợ Bến Thành khánh thành năm 1914, nổi tiếng với tháp đồng hồ bốn mặt ở Cửa Nam. Bên trong là không gian giao thương sầm uất với hàng ngàn gian hàng ẩm thực Nam Bộ, cà phê, lụa tơ tằm và đồ thủ công mỹ nghệ đậm đà bản sắc phương Nam.",
        cover_img="images/ben_thanh_main.jpg",
        grammar_str="Participle clause (Opened in 1914), Modal advice (You should practice friendly bargaining)"
    )

    # =========================================================================
    # SLIDE 7: NHÓM 4 - YOUNG PIONEERS
    # =========================================================================
    build_group_slide(
        prs,
        group_num=4,
        group_name="Young Pioneers",
        site_vn="Bến Nhà Rồng - Bảo tàng Hồ Chí Minh",
        site_en="Dragon Wharf - Ho Chi Minh Museum",
        duration_str="55 giây",
        words_num="92",
        script_en="Welcome to Dragon Wharf, officially known as the Ho Chi Minh Museum, situated along the Saigon River in District 4. Originally built in 1863 as a commercial port headquarters, the rooftop features two distinct ceramic dragon statues. The site holds tremendous historical significance. On June 5th, 1911, young patriot Nguyen Tat Thanh set sail abroad to find a path for national liberation. Today, it preserves thousands of precious artifacts, documents, and photographs of President Ho Chi Minh's life. Please visit respectfully and explore our history.",
        script_vn="Bến Nhà Rồng xây dựng năm 1863 bên sông Sài Gòn, nổi bật với đôi rồng gốm men xanh trên nóc mái. Đây là di tích lịch sử thiêng liêng - nơi người thanh niên yêu nước Nguyễn Tất Thành ra đi tìm đường cứu nước ngày 5/6/1911.",
        cover_img="images/dragon_wharf_main.jpg",
        grammar_str="Passive past (built in 1863), Historical time marker (On June 5th, 1911)"
    )

    # =========================================================================
    # SLIDE 8: TIẾT 2 - BÁO CÁO, KIỂM DUYỆT & TƯƠNG TÁC SỐ TẠI LỚP
    # =========================================================================
    s8 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s8, LIGHT_BG)
    add_slide_header(s8, "TIẾT 2: BÁO CÁO THUYẾT TRÌNH, KIỂM DUYỆT & TƯƠNG TÁC SỐ")

    card_left = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.35), Inches(6.0), Inches(5.5))
    card_left.fill.solid()
    card_left.fill.fore_color.rgb = WHITE
    card_left.line.color.rgb = NAVY
    card_left.line.width = Pt(1.5)
    tf_cl = card_left.text_frame
    tf_cl.word_wrap = True

    p = tf_cl.paragraphs[0]
    p.text = "🎯 CÔNG NGHỆ DẠY HỌC SỐ ĐÃ ÁP DỤNG TẠI LỚP 8A2:"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY

    items_s8 = [
        "1. Chế độ trình chiếu Poster số (Presentation Mode): Mở toàn màn hình, học sinh thuyết trình tiếng Anh tự tin kết hợp phát đoạn thu âm chuẩn.",
        "2. Bảng kiểm duyệt của Giáo viên (Teacher Verification Modal): Giáo viên trực tiếp kiểm tra 80 - 110 từ, duyệt hình ảnh và bấm '✓ Thông tin đã được phê duyệt' minh bạch.",
        "3. Tương tác số đa chiều (Mobile QR Interaction): Học sinh dùng điện thoại hoặc máy tính bảng quét mã QR trên màn hình chiếu.",
        "4. Đánh giá nghe hiểu nhanh: Trắc nghiệm 3 câu hỏi tiếng Anh cho mỗi trạm, hệ thống chấm điểm tự động tức thì.",
        "5. Thả tim & Bình luận chéo: Các nhóm quét mã QR để lại phản hồi, khen ngợi và góp ý phát âm tiếng Anh cho nhóm bạn."
    ]
    for itm in items_s8:
        p = tf_cl.add_paragraph()
        p.text = itm + "\n"
        p.font.size = Pt(11)
        p.font.color.rgb = DARK

    # Right side: 2 Placeholders for Classroom Photos
    ph_w = Inches(5.8)
    ph_h = Inches(2.6)

    h1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.35), ph_w, ph_h)
    h1.fill.solid()
    h1.fill.fore_color.rgb = RGBColor(238, 242, 255) # soft indigo
    h1.line.color.rgb = RGBColor(99, 102, 241)
    h1.line.width = Pt(2)
    tf_h1 = h1.text_frame
    tf_h1.word_wrap = True
    tf_h1.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_h1.paragraphs[0]
    p.text = "📷 VỊ TRÍ CHÈN ẢNH HỌC SINH THUYẾT TRÌNH TẠI LỚP TIẾT 2"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(67, 56, 202)
    p.alignment = PP_ALIGN.CENTER
    p = tf_h1.add_paragraph()
    p.text = "(Ảnh đại diện nhóm đứng trước lớp thuyết trình tiếng Anh trên màn chiếu)\n👉 Thầy/Cô gửi ảnh để em chèn vào đây."
    p.font.size = Pt(10.5)
    p.font.color.rgb = DARK
    p.alignment = PP_ALIGN.CENTER

    h2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(4.2), ph_w, ph_h)
    h2.fill.solid()
    h2.fill.fore_color.rgb = RGBColor(236, 253, 245) # soft emerald
    h2.line.color.rgb = EMERALD
    h2.line.width = Pt(2)
    tf_h2 = h2.text_frame
    tf_h2.word_wrap = True
    tf_h2.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf_h2.paragraphs[0]
    p.text = "📷 VỊ TRÍ CHÈN ẢNH GIÁO VIÊN KIỂM DUYỆT & HS QUÉT MÃ QR"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = RGBColor(6, 95, 70)
    p.alignment = PP_ALIGN.CENTER
    p = tf_h2.add_paragraph()
    p.text = "(Ảnh Cô giáo thao tác kiểm duyệt trên máy tính và học sinh quét mã QR bằng điện thoại)\n👉 Thầy/Cô gửi ảnh để em chèn vào đây."
    p.font.size = Pt(10.5)
    p.font.color.rgb = DARK
    p.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 9: HIỆU QUẢ DỰ ÁN & GIÁ TRỊ NHÂN RỘNG
    # =========================================================================
    s9 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s9, LIGHT_BG)
    add_slide_header(s9, "KẾT QUẢ ĐẠT ĐƯỢC & Ý NGHĨA THỰC TIỄN CỦA DỰ ÁN")

    card_w3 = Inches(3.8)
    card_gap3 = Inches(0.3)
    y3 = Inches(1.4)
    h3 = Inches(5.4)

    results = [
        ("NĂNG LỰC TIẾNG ANH & KỸ NĂNG MỀM", EMERALD, [
            "100% học sinh 4 nhóm tham gia tích cực, cải thiện phát âm và ngữ điệu tiếng Anh.",
            "Nắm vững các cấu trúc ngữ pháp trọng tâm: Passive Voice, Modal Verbs, Participle Clauses.",
            "Rèn luyện kỹ năng làm việc nhóm chuẩn STEM, tự tin thuyết trình trước tập thể.",
            "Học sinh chuyển từ 'học thụ động' sang 'chủ động kiến tạo sản phẩm' phục vụ cộng đồng."
        ]),
        ("CHUYỂN ĐỔI SỐ & HỌC LIỆU MỞ", NAVY, [
            "Website https://english8-two.vercel.app/ hoạt động 24/7 trên hạ tầng đám mây hiện đại.",
            "Dung lượng tối ưu, tải nhanh trên điện thoại di động mà không cần cài app phức tạp.",
            "Tích hợp mã QR Code cho từng di tích, sẵn sàng in thành tờ gấp, thẻ đeo cho du khách.",
            "Kho học liệu số mở phục vụ tái sử dụng cho các tiết học ôn tập và ngoại khóa."
        ]),
        ("GIÁ TRỊ THỰC TIỄN & NHÂN RỘNG", AMBER, [
            "Quảng bá hình ảnh văn hóa, lịch sử TP.HCM tươi đẹp, thân thiện đến bạn bè quốc tế.",
            "Mô hình dễ dàng chuyển giao cho các khối lớp 6, 7, 9 và các môn Lịch sử - Địa lý, GDĐP.",
            "Khẳng định tinh thần đổi mới phương pháp giảng dạy theo định hướng Chương trình GDPT 2018.",
            "Sản phẩm số thực chất, kết hợp học đường với đời sống thực tiễn sinh động."
        ])
    ]

    for idx, (title, col, items) in enumerate(results):
        x = Inches(0.6) + idx * (card_w3 + card_gap3)
        c = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y3, card_w3, h3)
        c.fill.solid()
        c.fill.fore_color.rgb = WHITE
        c.line.color.rgb = col
        c.line.width = Pt(2)

        # Header tag
        tag = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.15), y3 + Inches(0.15), card_w3 - Inches(0.3), Inches(0.7))
        tag.fill.solid()
        tag.fill.fore_color.rgb = col
        tag.line.fill.background()
        tf = tag.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

        # Body items
        body = s9.shapes.add_textbox(x + Inches(0.15), y3 + Inches(0.95), card_w3 - Inches(0.3), Inches(4.2))
        b_tf = body.text_frame
        b_tf.word_wrap = True
        for i_idx, item in enumerate(items):
            p = b_tf.paragraphs[0] if i_idx == 0 else b_tf.add_paragraph()
            p.text = f"✔ {item}\n"
            p.font.size = Pt(11)
            p.font.color.rgb = DARK

    # =========================================================================
    # SLIDE 10: KẾT LUẬN & TRẢI NGHIỆM SẢN PHẨM SỐ
    # =========================================================================
    s10 = prs.slides.add_slide(blank_slide_layout)
    set_slide_background(s10, NAVY)

    # Top border bar
    bar = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.18))
    bar.fill.solid()
    bar.fill.fore_color.rgb = GOLD
    bar.line.fill.background()

    # Content Box
    c_box = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.0), Inches(10.33), Inches(5.5))
    c_box.fill.solid()
    c_box.fill.fore_color.rgb = RGBColor(10, 56, 92)
    c_box.line.color.rgb = GOLD
    c_box.line.width = Pt(2)

    tf_c = c_box.text_frame
    tf_c.word_wrap = True
    tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf_c.paragraphs[0]
    p.text = "TRÂN TRỌNG CẢM ƠN QUÝ BAN GIÁM KHẢO!"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.alignment = PP_ALIGN.CENTER

    p = tf_c.add_paragraph()
    p.text = "DỰ ÁN HƯỚNG DẪN VIÊN SỐ DI SẢN SÀI GÒN BẰNG TIẾNG ANH (LỚP 8A2)\nTRƯỜNG THCS HÒA HIỆP - TP. HỒ CHÍ MINH"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    p = tf_c.add_paragraph()
    p.text = "\n🌐 TRẢI NGHIỆM TRỰC TIẾP HỆ THỐNG SẢN PHẨM SỐ TẠI:"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = RGBColor(254, 240, 138)
    p.alignment = PP_ALIGN.CENTER

    p = tf_c.add_paragraph()
    p.text = "👉 https://english8-two.vercel.app/ 👈"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER

    p = tf_c.add_paragraph()
    p.text = "\nQuét mã QR Code trên màn hình hoặc truy cập đường link để nghe toàn bộ 4 bản Audio Guide,\nxem Poster tương tác và tham gia kiểm duyệt trực tiếp cùng cô trò Trường THCS Hòa Hiệp!"
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(226, 232, 240)
    p.alignment = PP_ALIGN.CENTER

    # Save to both folders
    p1 = '01_Ho_So_Du_Thi_SGDDT/Bao_Cao_Thuyet_Trinh_STEM_Lop8.pptx'
    p2 = '02_Bao_Cao_Demo_Gui_Khach/Bao_Cao_Thuyet_Trinh_STEM_Lop8.pptx'
    prs.save(p1)
    prs.save(p2)
    print(f"Saved PowerPoint presentation to:\n- {p1}\n- {p2}")

if __name__ == '__main__':
    create_deck()
