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

def build_part_1_grade8(filename="Tienganh_THCS_Ho-Va-Ten_STEM_P1.docx"):
    doc = docx.Document()
    setup_document_styles(doc)

    # Header Table
    tbl_header = doc.add_table(rows=1, cols=2)
    tbl_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_header.autofit = False

    c0, c1 = tbl_header.rows[0].cells[0], tbl_header.rows[0].cells[1]
    c0.width = Inches(3.2); c1.width = Inches(3.5)

    p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p0.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO\nTHÀNH PHỐ HỒ CHÍ MINH\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p0.add_run("HỘI THI THIẾT KẾ CHỦ ĐỀ STEM/STEAM,\nCHỦ ĐỀ DẠY HỌC SỐ NĂM HỌC 2026 - 2027"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p1.add_run("Độc lập - Tự do - Hạnh phúc\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p1.add_run("-----------------------"); r.font.size = Pt(10)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(10); p_title.paragraph_format.space_after = Pt(2)
    r = p_title.add_run("PHẦN 1: PHẦN TỔNG QUAN\n"); r.font.bold = True; r.font.size = Pt(15); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_title.add_run("THÔNG TIN THÍ SINH DỰ THI VÀ BỐI CẢNH THỰC HIỆN CHỦ ĐỀ\n"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p_title.add_run("MÔN TIẾNG ANH – CẤP TRUNG HỌC CƠ SỞ (LỚP 8 – BỘ SÁCH GLOBAL SUCCESS)"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r = p_sub.add_run("(Thực hiện theo Kế hoạch số 12070/KH-SGDĐT ngày 09/11/2026 của Sở GD&ĐT TP. Hồ Chí Minh)"); r.font.italic = True; r.font.size = Pt(10.5)

    # I. THÔNG TIN GIỚI THIỆU
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8); h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("I. THÔNG TIN GIỚI THIỆU VỀ GIÁO VIÊN VÀ ĐƠN VỊ"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    tbl_info = doc.add_table(rows=7, cols=2)
    tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_info.autofit = False
    set_table_borders(tbl_info, color="B0C4DE", sz="4")

    info_data = [
        ("Họ và tên giáo viên:", "[Họ và tên giáo viên dự thi]"),
        ("Năm sinh / Giới tính:", "[Năm sinh]   |   Giới tính: [Nam/Nữ]"),
        ("Bộ môn đăng ký dự thi:", "Tiếng Anh (Cấp Trung học cơ sở - Lớp 8)"),
        ("Thâm niên công tác:", "[Số năm] năm giảng dạy môn Tiếng Anh THCS"),
        ("Tên đơn vị công tác:", "Trường Trung học Cơ sở [Tên trường THCS], TP. Hồ Chí Minh"),
        ("Số điện thoại liên lạc cá nhân:", "[Số điện thoại di động giáo viên]"),
        ("Địa chỉ thư điện tử (Email):", "[Email giáo viên]@gmail.com / [Email giáo viên]@hcm.edu.vn")
    ]

    for idx, (lbl, val) in enumerate(info_data):
        row = tbl_info.rows[idx]
        c_l, c_v = row.cells[0], row.cells[1]
        c_l.width = Inches(2.3); c_v.width = Inches(4.4)
        set_cell_margins(c_l, 70, 70, 90, 90); set_cell_margins(c_v, 70, 70, 90, 90)
        set_cell_background(c_l, "F4F7FA")

        p = c_l.paragraphs[0]; r = p.add_run(lbl); r.font.bold = True; r.font.size = Pt(11.5)
        p = c_v.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(11.5)

    # II. TÊN VÀ TÓM LƯỢC NỘI DUNG CHỦ ĐỀ
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(12); h2.paragraph_format.space_after = Pt(4)
    r = h2.add_run("II. TÊN VÀ TÓM LƯỢC NỘI DUNG CHỦ ĐỀ DỰ THI"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Topic Box
    tbl_name = doc.add_table(rows=1, cols=1)
    tbl_name.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_name.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_name, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_name.rows[0].cells[0].paragraphs[0]; p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_box.add_run("TÊN CHỦ ĐỀ DẠY HỌC STEM:\n"); r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    r = p_box.add_run("SAIGON HERITAGE QR AUDIO GUIDE – VOICES OF OUR COMMUNITY\n"); r.font.bold = True; r.font.size = Pt(13.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_box.add_run("(Thiết kế bảng thuyết minh thông minh tích hợp mã QR hướng dẫn âm thanh tiếng Anh giới thiệu di tích lịch sử - văn hóa TP.HCM dành cho học sinh Lớp 8)"); r.font.italic = True; r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run("1. Đơn vị môn học đăng ký dự thi: "); r.font.bold = True
    p.add_run("Môn Tiếng Anh – Cấp THCS (Lớp 8 – Bộ sách Tiếng Anh 8 Global Success).\n")

    r = p.add_run("2. Tình huống thực tiễn và vấn đề đặt ra:\n"); r.font.bold = True
    p.add_run(
        "Thành phố Hồ Chí Minh là trung tâm văn hóa và du lịch sôi động, chào đón hàng triệu du khách quốc tế mỗi năm đến thăm các địa danh nổi tiếng "
        "như Bưu điện Trung tâm Sài Gòn, Nhà hát Thành phố, Bến Nhà Rồng, Chợ Bến Thành, v.v. "
        "Tuy nhiên, nhiều bảng thông tin tại các điểm di tích hiện nay chỉ viết bằng tiếng Việt hoặc câu chữ quá dài, thiếu tính tương tác sinh động. "
        "Du khách nước ngoài (đặc biệt là khách du lịch tự túc và các bạn trẻ quốc tế) khó tiếp cận nhanh câu chuyện lịch sử thú vị của địa điểm trong thời gian ngắn."
    )

    p_q = doc.add_paragraph()
    p_q.paragraph_format.left_indent = Inches(0.4)
    r = p_q.add_run("Câu hỏi dẫn dắt dự án (Driving Question):\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p_q.add_run("“Làm thế nào để học sinh lớp 8 có thể trở thành 'Đại sứ văn hóa nhí', giúp một du khách nước ngoài hiểu được câu chuyện đặc sắc của di tích TP.HCM chỉ trong 60 giây bằng tiếng Anh, thông qua một thao tác quét mã QR tiện lợi ngay trên điện thoại thông minh?”")
    r.font.bold = True; r.font.italic = True

    p = doc.add_paragraph()
    r = p.add_run("3. Tóm lược nội dung chủ đề:\n"); r.font.bold = True
    p.add_run(
        "Chủ đề được thiết kế trong thời lượng 02 tiết học trên lớp kết hợp hoạt động trải nghiệm ngoài lớp, "
        "bám sát kiến thức ngôn ngữ của sách Tiếng Anh 8 Global Success (Unit 5: Our Customs and Traditions & Unit 11: Science and Technology). "
        "Học sinh làm việc theo 4 nhóm học sinh (mỗi nhóm 8 - 10 em phụ trách 1 di tích), đóng vai các 'Nhà sáng tạo nội dung số nhí' (Junior Digital Creators). "
        "Các em vận dụng kỹ năng viết tiếng Anh súc tích (kịch bản 80 - 110 từ), cấu trúc ngữ pháp chỉ dẫn và miêu tả (was built, is famous for, you should...), "
        "kết hợp các nguyên lý Khoa học tự nhiên về âm thanh (tiêu âm, giảm tiếng ồn), ứng dụng Công nghệ số (thu âm smartphone, tạo mã QR, Canva), "
        "quy trình Kỹ thuật (thiết kế mẫu → thử nghiệm Gallery Walk trong lớp → cải tiến) và Toán học thống kê (tính tỷ lệ % quét mã, % hiểu bài).\n"
        "Sản phẩm cuối cùng của mỗi nhóm là 01 Bảng thuyết minh di sản khổ A4 tích hợp mã QR dẫn đến tệp Audio Guide tiếng Anh 45 - 60 giây đạt chuẩn âm thanh trong trẻo."
    )

    p = doc.add_paragraph()
    r = p.add_run("4. Điểm mới, độc đáo và phù hợp lứa tuổi học sinh lớp 8:\n"); r.font.bold = True
    p.add_run(
        "• Vừa sức và tạo hứng thú cao: Thay vì viết các bài văn dài hàn lâm, học sinh lớp 8 được rèn luyện kỹ năng nói - thu âm thực tế (Micro-podcasting 60 giây), giúp các em vượt qua tâm lý ngại nói tiếng Anh.\n"
        "• Thiết kế hai tầng thông tin (Dual-layer Experience): Tầng 1 là Audio Guide nhanh 60 giây; Tầng 2 là bảng in A4 có hình ảnh sinh động và đoạn tóm tắt chữ lớn (hỗ trợ người khiếm thính nhẹ hoặc người cao tuổi).\n"
        "• Thực hành đo lường bằng con số thực tế (Evidence-based STEM): Học sinh đóng vai du khách quét mã, làm bài trắc nghiệm nhanh 3 câu và tính toán số liệu thống kê (tỷ lệ quét thành công ≥ 90%, tỷ lệ nghe hiểu ≥ 80%) để hoàn thành phiếu cải tiến kỹ thuật.\n"
        "• Đóng góp vào Kho học liệu số TP.HCM: Xây dựng Website Cổng học liệu số Di sản Sài Gòn (Saigon Heritage Digital Portal): Nơi lưu trữ đám mây trực quan, cho phép học sinh tải mã QR, kịch bản song ngữ và tích hợp sẵn trình phát âm thanh (Audio Player) trực tiếp để du khách và cộng đồng truy cập mọi lúc mọi nơi."
    )

    # III. VỊ TRÍ KIẾN THỨC VÀ TÍCH HỢP STEM LỚP 8
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(12); h3.paragraph_format.space_after = Pt(4)
    r = h3.add_run("III. VỊ TRÍ KIẾN THỨC VÀ TÍCH HỢP LIÊN MÔN STEM/STEAM CHO LỚP 8"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Vị trí trong Chương trình Tiếng Anh 8 Global Success (GDPT 2018):\n"); r.font.bold = True
    p.add_run(
        "• Bám sát trực tiếp: Unit 5: Our Customs and Traditions (Phần Looking Back & Project) kết hợp với Unit 11: Science and Technology.\n"
        "• Kiến thức ngôn ngữ trọng tâm: Từ vựng về phong tục, di sản, truyền thống và công nghệ số; Cấu trúc câu bị động cơ bản thì quá khứ đơn (was/were built, was designed); Động từ khuyết thiếu đưa ra lời khuyên cho du khách (should / shouldn't); Mạo từ (a/an/the) và cấu trúc miêu tả địa điểm nổi tiếng (be famous for, be located in)."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Xác định Kiến thức và Kỹ năng nền (Prerequisites):\n"); r.font.bold = True
    p.add_run(
        "• Tiếng Anh: Vốn từ vựng cơ bản lớp 6, 7 về quê hương, địa điểm du lịch; thì quá khứ đơn (Past Simple).\n"
        "• Khoa học tự nhiên (Lớp 7, 8): Khái niệm cơ bản về nguồn âm, độ to của âm (dB), tiếng ồn và ô nhiễm tiếng ồn, các vật liệu hấp thụ âm tốt (vải, rèm, xốp).\n"
        "• Tin học / Công nghệ (Lớp 8): Kỹ năng tìm kiếm thông tin trên Internet, sử dụng điện thoại thông minh để ghi âm, chia sẻ liên kết trên Google Drive.\n"
        "• Toán học (Lớp 8): Thu thập số liệu, tính tỷ số phần trăm (%), tính trung bình cộng và biểu diễn bảng số liệu."
    )

    p = doc.add_paragraph()
    r = p.add_run("3. Ma trận tích hợp liên môn STEM/STEAM lớp 8:\n"); r.font.bold = True

    tbl_stem = doc.add_table(rows=6, cols=3)
    tbl_stem.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_stem.autofit = False
    set_table_borders(tbl_stem, color="B0C4DE", sz="4")

    stem_data = [
        ("Thành tố STEM", "Kiến thức môn học lớp 8", "Biểu hiện cụ thể trong chủ đề bài học"),
        ("S\n(Science - Khoa học)", "KHTN 7-8: Chủ đề Âm thanh (sóng âm, nguồn âm, tiếng ồn, vật liệu tiêu âm).", "Học sinh hiểu được nguyên nhân làm rè tiếng và tạp âm; tìm không gian kín có rèm/vải để thu âm; giữ mic điện thoại cách miệng 1 gang tay (15 cm) chếch 45 độ để tránh luồng thở mạnh làm vỡ tiếng."),
        ("T\n(Technology - Công nghệ)", "Tin học 8: Sử dụng phần mềm đa phương tiện, lưu trữ đám mây, công cụ số.", "Sử dụng ứng dụng ghi âm trên smartphone, cắt gọt đoạn thừa, tải tệp lên Google Drive/Canva, tạo mã QR bằng web miễn phí (me-qr.com/qr-code-generator)."),
        ("E\n(Engineering - Kỹ thuật)", "Quy trình thiết kế kỹ thuật (EDP) 5 bước phù hợp cấp THCS.", "Thực hiện chu trình: 1. Nêu vấn đề → 2. Lập kịch bản 5 bước & vẽ mẫu bảng → 3. Chế tạo bảng mẫu có mã QR → 4. Thử nghiệm Gallery Walk → 5. Đánh giá và cải tiến."),
        ("A\n(Art - Nghệ thuật)", "Mỹ thuật 8: Thiết kế đồ họa trực quan, bố cục trang trí bảng A4 sinh động.", "Trang trí bảng thuyết minh A4 hài hòa, chọn ảnh di tích đẹp, tạo độ tương phản màu để mã QR dễ quét; rèn luyện giọng đọc tiếng Anh truyền cảm, ngữ điệu thân thiện."),
        ("M\n(Mathematics - Toán học)", "Toán 8: Thu thập và tổ chức dữ liệu, tính tỷ số phần trăm (%).", "Đo lường số liệu thử nghiệm: Tính tỷ lệ quét QR thành công (≥ 90%), tỷ lệ trả lời đúng câu hỏi nghe hiểu (≥ 80%), tính điểm trung bình độ rõ âm thanh (thang 1 - 5).")
    ]

    for idx, (c0_t, c1_t, c2_t) in enumerate(stem_data):
        row = tbl_stem.rows[idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(1.5); c1.width = Inches(2.3); c2.width = Inches(2.9)
        set_cell_margins(c0, 70, 70, 70, 70); set_cell_margins(c1, 70, 70, 70, 70); set_cell_margins(c2, 70, 70, 70, 70)

        if idx == 0:
            for c in [c0, c1, c2]: set_cell_background(c, "003366")
            p = c0.paragraphs[0]; r = p.add_run(c0_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c1.paragraphs[0]; r = p.add_run(c1_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c2.paragraphs[0]; r = p.add_run(c2_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            set_cell_background(c0, "F4F7FA")
            p = c0.paragraphs[0]; r = p.add_run(c0_t); r.font.bold = True; r.font.size = Pt(11)
            p = c1.paragraphs[0]; r = p.add_run(c1_t); r.font.size = Pt(11)
            p = c2.paragraphs[0]; r = p.add_run(c2_t); r.font.size = Pt(11)

    # IV. PHẨM CHẤT VÀ NĂNG LỰC
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(12); h4.paragraph_format.space_after = Pt(4)
    r = h4.add_run("IV. PHẨM CHẤT VÀ NĂNG LỰC HỌC SINH ĐẠT ĐƯỢC (THEO CT GDPT 2018)"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Năng lực đặc thù – Năng lực Giao tiếp Tiếng Anh lớp 8:\n"); r.font.bold = True
    p.add_run(
        "• Kỹ năng Nghe: Nghe hiểu các bài giới thiệu di tích đơn giản; nghe phát hiện thông tin chi tiết và nhận diện các lỗi phát âm/ngữ điệu khi đi thử nghiệm trạm của nhóm bạn.\n"
        "• Kỹ năng Nói: Phát âm rõ ràng các âm cuối (/s/, /z/, /ed/), nhấn đúng trọng âm từ; nói đoạn văn 45 - 60 giây trôi chảy với ngữ điệu tự nhiên, thân thiện; tự tin giới thiệu giải pháp nhóm.\n"
        "• Kỹ năng Đọc: Đọc quét thông tin lịch sử di tích TP.HCM từ các nguồn tư liệu học tập định hướng của giáo viên.\n"
        "• Kỹ năng Viết: Viết kịch bản ngắn gọn (80 - 110 từ) theo khung 5 bước: Hook → Location → Architecture/History → Fun fact → Advice/Invitation, sử dụng đúng thì quá khứ đơn, câu bị động cơ bản và động từ khuyết thiếu should."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Năng lực chung:\n"); r.font.bold = True
    p.add_run(
        "• Tự chủ và tự học: Tự giác tìm tài liệu, chủ động luyện đọc kịch bản và thử nghiệm thu âm nhiều lần.\n"
        "• Giao tiếp và hợp tác: Biết lắng nghe, phân công công việc công bằng trong nhóm 4 - 5 học sinh; đóng góp ý kiến nhận xét chân thành cho nhóm bạn.\n"
        "• Giải quyết vấn đề và sáng tạo: Biết ứng dụng công nghệ mã QR và âm thanh số để giải quyết bài toán giao tiếp thực tế cho du khách."
    )

    p = doc.add_paragraph()
    r = p.add_run("3. Năng lực số:\n"); r.font.bold = True
    p.add_run("• Biết sử dụng các ứng dụng ghi âm trên điện thoại, trang web tạo mã QR, Google Drive và Canva; có ý thức bảo vệ tài khoản và tôn trọng bản quyền hình ảnh.")

    p = doc.add_paragraph()
    r = p.add_run("4. Phẩm chất chủ yếu:\n"); r.font.bold = True
    p.add_run(
        "• Yêu nước: Yêu quý và tự hào về các di tích văn hóa, kiến trúc lịch sử của TP. Hồ Chí Minh; mong muốn giới thiệu nét đẹp thành phố với bạn bè năm châu.\n"
        "• Trách nhiệm: Có ý thức bảo vệ di sản văn hóa, giữ gìn vệ sinh môi trường tại các điểm du lịch; có trách nhiệm hoàn thành nhiệm vụ nhóm.\n"
        "• Chăm chỉ: Kiên nhẫn luyện phát âm, chăm chút từng chi tiết trình bày trên bảng thuyết minh."
    )

    # V. THỜI LƯỢNG VÀ KHÔNG GIAN
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(12); h5.paragraph_format.space_after = Pt(4)
    r = h5.add_run("V. THỜI LƯỢNG VÀ KHÔNG GIAN THỰC HIỆN"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Thời lượng thực hiện: "); r.font.bold = True
    p.add_run("Đúng chuẩn quy định ")
    r_b = p.add_run("02 tiết trên lớp (90 phút)"); r_b.font.bold = True
    p.add_run(", kết hợp với các hoạt động nghiên cứu và chế tạo ngoài giờ học theo mô hình lớp học đảo ngược:")

    tbl_time = doc.add_table(rows=5, cols=3)
    tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_time.autofit = False
    set_table_borders(tbl_time, color="B0C4DE", sz="4")

    time_data = [
        ("Giai đoạn thực hiện", "Nội dung hoạt động trọng tâm", "Thời lượng"),
        ("Giai đoạn 1\n(Trước Tiết 1 - Ngoài lớp)", "Chọn 1 di tích TP.HCM (trong danh sách gợi ý của GV), tìm hiểu 3 thông tin nổi bật, viết bản nháp kịch bản tiếng Anh 80 từ.", "Giao trước 3 - 5 ngày"),
        ("Giai đoạn 2\n(Tiết 1 - Trên lớp học)", "Khám phá vấn đề thực tiễn → Phân tích audio mẫu → Thống nhất tiêu chí kỹ thuật → Hoàn thiện kịch bản tiếng Anh & vẽ phác thảo bố cục bảng A4.", "45 phút (Tiết 1)"),
        ("Giai đoạn 3\n(Giữa 2 tiết - Ngoài lớp)", "Thu âm tiếng Anh tại nhà (cự ly 1 gang tay, phòng yên tĩnh) → Tải lên Drive → Tạo mã QR → Thiết kế bảng A4 trên Canva và in màu.", "1 - 2 ngày giữa hai tiết"),
        ("Giai đoạn 4\n(Tiết 2 - Trên lớp học)", "Triển lãm Gallery Walk (đóng vai du khách quét mã, nghe audio, làm trắc nghiệm) → Thu thập và tính toán số liệu → Cải tiến kỹ thuật → Tổng kết.", "45 phút (Tiết 2)")
    ]

    for idx, (t1, t2, t3) in enumerate(time_data):
        row = tbl_time.rows[idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(1.8); c1.width = Inches(3.4); c2.width = Inches(1.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60); set_cell_margins(c2, 60, 60, 60, 60)

        if idx == 0:
            for c in [c0, c1, c2]: set_cell_background(c, "003366")
            p = c0.paragraphs[0]; r = p.add_run(t1); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c1.paragraphs[0]; r = p.add_run(t2); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c2.paragraphs[0]; r = p.add_run(t3); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            set_cell_background(c0, "F4F7FA")
            p = c0.paragraphs[0]; r = p.add_run(t1); r.font.bold = True; r.font.size = Pt(11)
            p = c1.paragraphs[0]; r = p.add_run(t2); r.font.size = Pt(11)
            p = c2.paragraphs[0]; r = p.add_run(t3); r.font.size = Pt(11)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run("2. Không gian thực hiện:\n"); r.font.bold = True
    p.add_run(
        "• Trong lớp học: Bố trí không gian tương tác nhóm và 4 trạm di tích (Bưu điện Trung tâm, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) trưng bày quanh lớp cho hoạt động triển lãm Gallery Walk.\n"
        "• Ngoài lớp học: Góc học tập yên tĩnh tại nhà để thu âm; tham quan thực tế di tích hoặc xem ảnh 360 độ trên Internet.\n"
        "• Không gian số: Website Cổng học liệu số Di sản Sài Gòn (Saigon Heritage Digital Portal) do GV thiết kế riêng cho dự án, kết hợp Google Drive lưu trữ âm thanh."
    )

    # VI. MỤC TIÊU CẦN ĐẠT ĐƯỢC
    h6 = doc.add_paragraph()
    h6.paragraph_format.space_before = Pt(12); h6.paragraph_format.space_after = Pt(4)
    r = h6.add_run("VI. MỤC TIÊU CẦN ĐẠT ĐƯỢC SAU CHỦ ĐỀ"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    p.add_run(
        "1. Kiến thức tiếng Anh: 100% học sinh lớp 8 nắm vững và sử dụng đúng từ vựng di tích, cấu trúc bị động thì quá khứ đơn (was built, was opened), cấu trúc be famous for và động từ khuyết thiếu should khi hướng dẫn du khách.\n"
        "2. Sản phẩm học tập: Mỗi nhóm hoàn thành 01 Bộ sản phẩm số gồm: 1 Kịch bản tiếng Anh 5 bước, 1 Audio Guide 45 - 60s, 1 mã QR và Bộ 5 ảnh chụp di tích (1 ảnh chính toàn cảnh + 4 ảnh phụ chi tiết kiến trúc/nội thất) được xuất bản trên Website tiếng Anh 45 - 60 giây, đạt cự ly quét 20 - 30 cm trong vòng 3 giây với tỷ lệ thành công ≥ 90%.\n"
        "3. Kỹ năng STEM & Số: Nắm vững các bước thiết kế và cải tiến dựa trên số liệu thực tế đo được; sử dụng thành thạo các công cụ số cơ bản.\n"
        "4. Giá trị lan tỏa: Toàn bộ sản phẩm được tích hợp hoàn chỉnh trên Website Cổng học liệu số Di sản Sài Gòn, sẵn sàng liên kết và đóng góp vào Hệ thống quản lý học liệu số dùng chung ngành GDĐT TP. Hồ Chí Minh."
    )

    # Signature Block
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sign.autofit = False
    tbl_sign.rows[0].cells[0].width = Inches(3.2); tbl_sign.rows[0].cells[1].width = Inches(3.5)

    c0, c1 = tbl_sign.rows[0].cells[0], tbl_sign.rows[0].cells[1]
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("XÁC NHẬN CỦA ĐƠN VỊ TRƯỜNG THCS\n"); r.font.bold = True
    r = p.add_run("(Ký tên, đóng dấu)"); r.font.italic = True

    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TP. Hồ Chí Minh, ngày       tháng 11 năm 2026\n"); r.font.italic = True
    r = p.add_run("GIÁO VIÊN XÂY DỰNG CHỦ ĐỀ\n"); r.font.bold = True
    r = p.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"); r.font.italic = True
    r = p.add_run("[Họ và tên giáo viên]"); r.font.bold = True

    doc.save(filename)
    print(f"File {filename} created successfully.")

if __name__ == "__main__":
    build_part_1_grade8()
