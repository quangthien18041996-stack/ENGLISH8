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

def build_part_2_grade8(filename="Tienganh_THCS_Ho-Va-Ten_STEM_P2.docx"):
    doc = docx.Document()
    setup_document_styles(doc)

    # Header block
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_top.autofit = False
    c_l, c_r = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_l.width = Inches(3.2); c_r.width = Inches(3.5)

    p_l = c_l.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_l.add_run("Trường: THCS [Tên trường]\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p_l.add_run("Tổ bộ môn: Tiếng Anh"); r.font.bold = True; r.font.size = Pt(11)

    p_r = c_r.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_r.add_run("Họ và tên giáo viên:\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p_r.add_run("[Họ và tên giáo viên]"); r.font.size = Pt(11)

    # Document Reference
    p_ref = doc.add_paragraph()
    p_ref.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ref.paragraph_format.space_before = Pt(8); p_ref.paragraph_format.space_after = Pt(2)
    r = p_ref.add_run("Phụ lục IV – Khung Kế hoạch bài dạy (Công văn 5512/BGDĐT-GDTrH)\n")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    r = p_ref.add_run("HỘI THI THIẾT KẾ CHỦ ĐỀ STEM/STEAM VÀ CHỦ ĐỀ DẠY HỌC SỐ NĂM HỌC 2026 - 2027\n(Kèm theo Kế hoạch số 12070/KH-SGDĐT ngày 09/11/2026 của Sở GD&ĐT TP.HCM)")
    r.font.italic = True; r.font.size = Pt(10.5)

    # Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8); p_title.paragraph_format.space_after = Pt(4)
    r = p_title.add_run("KẾ HOẠCH BÀI DẠY CHỦ ĐỀ DẠY HỌC STEM\n")
    r.font.bold = True; r.font.size = Pt(15); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Callout Box for Lesson Title
    tbl_name = doc.add_table(rows=1, cols=1)
    tbl_name.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_name.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_name, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_name.rows[0].cells[0].paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_box.add_run("TÊN BÀI DẠY: ")
    r1.font.bold = True; r1.font.size = Pt(13); r1.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r2 = p_box.add_run("SAIGON HERITAGE QR AUDIO GUIDE – VOICES OF OUR COMMUNITY\n")
    r2.font.bold = True; r2.font.size = Pt(13.5); r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r3 = p_box.add_run("Môn học: Tiếng Anh   |   Lớp: 8 (Cấp THCS)   |   Thời gian thực hiện: 02 tiết (90 phút)\n")
    r3.font.bold = True; r3.font.size = Pt(11.5)
    r4 = p_box.add_run("(Tích hợp Unit 5: Our Customs and Traditions & Unit 11: Science and Technology – Tiếng Anh 8 Global Success)")
    r4.font.italic = True; r4.font.size = Pt(11)

    # I. MỤC TIÊU
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14); h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("I. MỤC TIÊU BÀI HỌC"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Về kiến thức:\n"); r.font.bold = True
    p.add_run(
        "• Môn Tiếng Anh 8 (Global Success):\n"
        "  - Từ vựng: Vận dụng thành thạo từ vựng về di sản, phong tục, địa điểm du lịch: heritage, tradition, historical monument, architecture, attract, preserve, audio guide, tourist...\n"
        "  - Ngữ pháp: Sử dụng chính xác câu bị động thì quá khứ đơn cơ bản (was/were built, was designed by...); động từ khuyết thiếu đưa ra lời khuyên cho khách du lịch (should / shouldn't); cấu trúc miêu tả địa điểm: be famous for, be located in, be known as...\n"
        "  - Cấu trúc kịch bản thuyết minh ngắn (Micro-podcast): Nắm vững 5 bước kịch bản 80 - 110 từ (Hook → Location → History/Architecture → Fun fact → Advice/Invitation).\n"
        "• Tích hợp STEM/STEAM lớp 8:\n"
        "  - Khoa học (S - KHTN 7-8): Hiểu về âm thanh, nguồn âm, tiếng ồn và ô nhiễm tiếng ồn; biết cách hạn chế tạp âm khi ghi âm bằng cách sử dụng vật liệu tiêu âm (rèm cửa, phòng kín) và giữ đúng cự ly micrô.\n"
        "  - Công nghệ (T - Tin học 8): Sử dụng ứng dụng ghi âm smartphone, chia sẻ tệp âm thanh công khai trên Google Drive, tạo mã QR bằng web miễn phí, thiết kế bảng thông tin trên Canva.\n"
        "  - Kỹ thuật (E): Vận dụng quy trình thiết kế kỹ thuật 5 bước (Nêu vấn đề → Thiết kế kịch bản & bảng mẫu → Chế tạo mẫu thử → Kiểm thử thực tế Gallery Walk → Cải tiến).\n"
        "  - Nghệ thuật (A - Mỹ thuật 8): Bố cục trang trí bảng A4 sinh động, tương phản màu sắc giúp mã QR nổi bật; rèn luyện giọng đọc diễn cảm, thân thiện.\n"
        "  - Toán học (M - Toán 8): Thu thập dữ liệu thực tế, tính tỷ lệ phần trăm (%) quét mã thành công (≥ 90%), tỷ lệ người nghe trả lời đúng nội dung (≥ 80%), tính điểm trung bình thang đo 1 - 5 sao."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Về năng lực:\n"); r.font.bold = True
    p.add_run(
        "a) Năng lực đặc thù môn Tiếng Anh lớp 8:\n"
        "  - Nghe: Nghe hiểu các thông tin chính về di tích TP.HCM; nghe phát hiện lỗi phát âm, tốc độ trong bài thử nghiệm của bạn.\n"
        "  - Nói: Phát âm chuẩn xác các âm đuôi (/s/, /z/, /ed/); nói đoạn thuyết minh 45 - 60 giây trôi chảy, tự nhiên; tự tin giới thiệu sản phẩm của nhóm.\n"
        "  - Đọc: Đọc quét thông tin lịch sử từ tư liệu do GV định hướng.\n"
        "  - Viết: Viết kịch bản tiếng Anh 80 - 110 từ đúng ngữ pháp, súc tích và giàu sức gợi.\n"
        "b) Năng lực chung:\n"
        "  - Tự chủ và tự học: Chủ động luyện đọc kịch bản, tự tìm hiểu ứng dụng công nghệ số.\n"
        "  - Giao tiếp và hợp tác: Hợp tác tích cực trong nhóm 4 - 5 em, biết phân công và hỗ trợ nhau hoàn thành sản phẩm.\n"
        "  - Giải quyết vấn đề và sáng tạo: Sáng tạo giải pháp số thông minh để quảng bá di tích quê hương đến du khách quốc tế.\n"
        "c) Năng lực số: Thao tác thuần thục điện thoại thông minh, mã QR, Google Drive và Canva an toàn, đúng quy định."
    )

    p = doc.add_paragraph()
    r = p.add_run("3. Về phẩm chất:\n"); r.font.bold = True
    p.add_run(
        "• Yêu nước: Yêu quý các di tích lịch sử, văn hóa TP.HCM; tự hào và muốn giới thiệu nét đẹp quê hương đến du khách bốn phương.\n"
        "• Trách nhiệm: Ý thức giữ gìn, bảo vệ di sản văn hóa; có tinh thần trách nhiệm với nhiệm vụ của nhóm.\n"
        "• Chăm chỉ: Tỉ mỉ luyện phát âm nhiều lần để có bản thu âm trong trẻo nhất; trang trí bảng thuyết minh cẩn thận."
    )

    # II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14); h2.paragraph_format.space_after = Pt(4)
    r = h2.add_run("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Thiết bị và học liệu của Giáo viên:\n"); r.font.bold = True
    p.add_run(
        "• Thiết bị: Máy tính, máy chiếu hoặc màn hình tương tác, loa, mạng Wi-Fi lớp học.\n"
        "• Học liệu số:\n"
        "  - Slide trình chiếu bài học tương tác (Canva / PowerPoint).\n"
        "  - Video ngắn (20 giây) về tình huống du khách nước ngoài tham quan di tích TP.HCM gặp rào cản ngôn ngữ.\n"
        "  - 02 file audio mẫu ngắn (45 giây): Audio 1 (đạt chuẩn) vs Audio 2 (nhiều tạp âm quạt gió, đọc quá nhanh, phát âm thiếu âm đuôi).\n"
        "  - Website Cổng học liệu số Di sản Sài Gòn (Saigon Heritage Digital Portal) do GV thiết kế, tích hợp bản đồ các di tích, trình phát âm thanh và cổng tiếp nhận sản phẩm của học sinh.\n"
        "  - Hệ thống Phiếu học tập số và Rubric đánh giá (Phụ lục 1, 2, 3, 4)."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Thiết bị và học liệu của Học sinh:\n"); r.font.bold = True
    p.add_run(
        "• Mỗi nhóm (4 - 5 học sinh): 01 - 02 điện thoại thông minh (smartphone) có camera quét mã QR và tai nghe cá nhân.\n"
        "• Dụng cụ: Giấy A4/A3, bút dạ màu, kéo, băng dính hai mặt để trang trí và trưng bày sản phẩm."
    )

    # III. TIẾN TRÌNH DẠY HỌC
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14); h3.paragraph_format.space_after = Pt(4)
    r = h3.add_run("III. TIẾN TRÌNH DẠY HỌC CHI TIẾT"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # ------------------ HOẠT ĐỘNG 1 ------------------
    h_act1 = doc.add_paragraph()
    h_act1.paragraph_format.space_before = Pt(12); h_act1.paragraph_format.space_after = Pt(4)
    r = h_act1.add_run("1. HOẠT ĐỘNG 1: XÁC ĐỊNH VẤN ĐỀ THỰC TIỄN / MỞ ĐẦU (Tiết 1 – 10 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act1.add_run("“RÀO CẢN NGÔN NGỮ CỦA DU KHÁCH VÀ DỰ ÁN ĐẠI SỨ VĂN HÓA NHÍ TP.HCM”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Tạo hứng thú, giúp học sinh lớp 8 nhận ra vấn đề thực tế tại các di tích TP.HCM; xác định nhiệm vụ thiết kế bảng thuyết minh A4 có mã QR Audio Guide tiếng Anh 45 - 60 giây và thống nhất 4 tiêu chí kỹ thuật sản phẩm.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run("Học sinh xem đoạn clip ngắn về du khách quốc tế trước bảng chỉ dẫn tiếng Việt tại Bưu điện TP.HCM; thảo luận cặp đôi trả lời Driving Question và làm rõ 4 tiêu chí kỹ thuật cần đạt.")

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run(
        "Bảng 4 Tiêu chí kỹ thuật sản phẩm (Design Specifications) được ghi vào vở:\n"
        "  1. Kịch bản tiếng Anh: 80 - 110 từ, đúng 5 phần, sử dụng thì quá khứ đơn, câu bị động và lời khuyên (should).\n"
        "  2. Tệp âm thanh Audio Guide: 45 - 60 giây, giọng đọc rõ ràng, tự nhiên, không bị tạp âm quạt gió hay tiếng thở rè.\n"
        "  3. Mã QR: Quét được dễ dàng từ cự ly 20 - 30 cm trong vòng 3 giây; link Google Drive mở công khai.\n"
        "  4. Bảng thông tin in A4: Thiết kế bắt mắt, có ảnh di tích, có tóm tắt chữ lớn cho người nghe kém."
    )

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act1 = doc.add_table(rows=4, cols=2)
    tbl_act1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act1.autofit = False
    set_table_borders(tbl_act1, color="B0C4DE", sz="4")

    act1_steps = [
        ("Bước 1: Giao nhiệm vụ (3 phút)", "• Giáo viên chiếu clip ngắn về du khách quốc tế tại di tích TP.HCM.\n• Đặt câu hỏi: 'Nếu em gặp du khách này, làm sao giúp họ hiểu câu chuyện di tích chỉ trong 1 phút bằng điện thoại?'\n• Giới thiệu dự án 'Saigon Heritage QR Audio Guide'."),
        ("Bước 2: Thực hiện nhiệm vụ (3 phút)", "• Học sinh làm việc theo cặp đôi (Think-Pair), thảo luận ý tưởng giải pháp.\n• Giáo viên gợi ý giải pháp âm thanh số qua mã QR tiện lợi trên điện thoại."),
        ("Bước 3: Báo cáo, thảo luận (2 phút)", "• Đại diện 2 học sinh chia sẻ ý tưởng: Bảng giấy truyền thống chữ nhỏ khó đọc, dùng mã QR nghe audio sẽ sinh động và nhanh chóng hơn."),
        ("Bước 4: Kết luận, nhận định (2 phút)", "• Giáo viên chốt lại bài toán và trình chiếu 4 tiêu chí kỹ thuật bắt buộc của sản phẩm mẫu.")
    ]

    for idx, (s_t, s_d) in enumerate(act1_steps):
        row = tbl_act1.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ------------------ HOẠT ĐỘNG 2 ------------------
    h_act2 = doc.add_paragraph()
    h_act2.paragraph_format.space_before = Pt(12); h_act2.paragraph_format.space_after = Pt(4)
    r = h_act2.add_run("2. HOẠT ĐỘNG 2: NGHIÊN CỨU KIẾN THỨC NỀN VÀ THIẾT KẾ GIẢI PHÁP (Tiết 1 – 35 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act2.add_run("“PHÂN TÍCH AUDIO MẪU, HỌC CẤU TRÚC VÀ THIẾT KẾ KỊCH BẢN - STORYBOARD BẢNG A4”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Học sinh phân biệt được audio đạt chuẩn và audio lỗi (tạp âm, tốc độ); nắm vững cấu trúc kịch bản 5 bước; viết hoàn chỉnh kịch bản tiếng Anh 80 - 110 từ và vẽ Storyboard bảng thuyết minh A4 của nhóm.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Nghe đối chiếu 2 đoạn audio mẫu 45 giây của GV, phát hiện lỗi về âm học (tiếng gió quạt, mic để quá gần miệng) và phát âm (nuốt âm cuối, đọc đều đều).\n"
        "• Ôn tập và áp dụng cấu trúc: was/were built in..., is famous for..., You should visit / take photos...\n"
        "• Viết kịch bản nhóm 5 bước: 1. Hook (Chào hỏi & câu hỏi 10s) → 2. Location (Vị trí 10s) → 3. History/Architecture (Năm xây dựng, nét đẹp 15s) → 4. Fun fact (1 điều thú vị 10s) → 5. Advice (Lời khuyên & lời chào 10s).\n"
        "• Biên tập chéo (Peer-editing) giữa các nhóm để đếm từ (80 - 110 từ) và sửa lỗi chính tả.\n"
        "• Vẽ phác thảo Storyboard bảng A4: vị trí dán mã QR (tối thiểu 4x4 cm), vị trí ảnh và dòng tóm tắt chữ lớn."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Bản kịch bản tiếng Anh 80 - 110 từ đã hoàn thiện và Bản vẽ Storyboard bảng A4 trên giấy của nhóm.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act2 = doc.add_table(rows=4, cols=2)
    tbl_act2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act2.autofit = False
    set_table_borders(tbl_act2, color="B0C4DE", sz="4")

    act2_steps = [
        ("Bước 1: Giao nhiệm vụ (5 phút)", "• Giáo viên mở 2 đoạn audio mẫu qua loa; phát Phiếu học tập số 1 (Phụ lục 1).\n• Hướng dẫn quy tắc âm thanh: Mic cách miệng 1 gang tay (15cm), nói chếch để không phà hơi vào mic.\n• Giao nhiệm vụ viết kịch bản 80 - 110 từ và vẽ Storyboard."),
        ("Bước 2: Thực hiện nhiệm vụ (20 phút)", "• Nhóm làm việc: Các thành viên cùng hoàn thiện kịch bản dựa trên di tích đã chọn; 1 bạn vẽ phác thảo bố cục bảng A4.\n• Đổi kịch bản kiểm tra chéo (Peer-edit) giữa các nhóm theo checklist.\n• Giáo viên đi quanh các nhóm hướng dẫn phát âm chuẩn các từ khó (architecture, monument, magnificent...)."),
        ("Bước 3: Báo cáo, thảo luận (6 phút)", "• 01 nhóm đọc thử kịch bản tiếng Anh trước lớp với ngữ điệu diễn cảm.\n• 01 nhóm trình bày bản vẽ Storyboard, giải thích vị trí dán mã QR và màu sắc trang trí.\n• Giáo viên và các bạn nhận xét, góp ý."),
        ("Bước 4: Kết luận, nhận định (4 phút)", "• Giáo viên chốt lại kịch bản chuẩn cho 4 nhóm.\n• Hướng dẫn chi tiết nhiệm vụ chế tạo ngoài lớp: Thu âm tại nhà, tải lên Drive, tạo mã QR và in màu bảng A4 chuẩn bị cho Tiết 2.")
    ]

    for idx, (s_t, s_d) in enumerate(act2_steps):
        row = tbl_act2.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ------------------ HOẠT ĐỘNG NGOÀI LỚP ------------------
    h_ext = doc.add_paragraph()
    h_ext.paragraph_format.space_before = Pt(12); h_ext.paragraph_format.space_after = Pt(4)
    r = h_ext.add_run("HOẠT ĐỘNG TRẢI NGHIỆM NGOÀI LỚP HỌC (GIỮA TIẾT 1 VÀ TIẾT 2)\n")
    r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x70, 0x30, 0xA0)
    r = h_ext.add_run("CHẾ TẠO SẢN PHẨM MẪU: THU ÂM TIẾNG ANH, TẠO MÃ QR VÀ IN ẤN BẢNG A4")
    r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    p_ext = doc.add_paragraph()
    p_ext.add_run(
        "Học sinh thực hiện tại nhà (1 - 2 ngày) dưới sự cố vấn trực tuyến của giáo viên:\n"
        "1. Thu âm tiếng Anh: Tìm phòng kín yên tĩnh (kéo rèm để chống vang âm); để điện thoại cách miệng 1 gang tay (15 cm); đọc kịch bản diễn cảm với thời lượng 45 - 60 giây.\n"
        "2. Lưu trữ đám mây & Mã QR: Tải file ghi âm lên Google Drive của nhóm, mở quyền xem công khai ('Bất kỳ ai có đường liên kết'); dùng trang web tạo mã QR trực tuyến miễn phí để tạo mã QR dẫn thẳng đến file nghe.\n"
        "3. Thiết kế bảng A4: Thiết kế bảng thông tin trên Canva (hoặc vẽ tay trang trí khổ A4), dán mã QR kích thước tối thiểu 4 x 4 cm, in màu mang đến lớp cho Tiết 2."
    )

    # ------------------ HOẠT ĐỘNG 3 ------------------
    h_act3 = doc.add_paragraph()
    h_act3.paragraph_format.space_before = Pt(12); h_act3.paragraph_format.space_after = Pt(4)
    r = h_act3.add_run("3. HOẠT ĐỘNG 3: THỬ NGHIỆM SẢN PHẨM, BÁO CÁO VÀ THẢO LUẬN (Tiết 2 – 25 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act3.add_run("“TRIỂN LÃM GALLERY WALK – TRẢI NGHIỆM QUÉT MÃ VÀ ĐO LƯỜNG THỰC TẾ”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Tổ chức cho học sinh lớp 8 trải nghiệm đóng vai du khách quốc tế; luyện kỹ năng nghe hiểu tiếng Anh thực tế; kiểm tra chất lượng kỹ thuật của bảng thuyết minh và thu thập số liệu đo đạc (tỷ lệ quét QR, tỷ lệ trả lời đúng câu hỏi, điểm âm thanh).")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• 4 nhóm dán bảng A4 tại 4 trạm quanh phòng học (Trạm 1: Bưu điện Trung tâm Sài Gòn; Trạm 2: Nhà thờ Đức Bà; Trạm 3: Chợ Bến Thành; Trạm 4: Bến Nhà Rồng).\n"
        "• 2 bạn ở lại trực trạm hướng dẫn và ghi chép số liệu; các bạn còn lại đeo tai nghe đi qua các trạm: Quét mã QR bằng điện thoại → Nghe audio guide tiếng Anh → Trả lời 3 câu hỏi trắc nghiệm nhanh trên phiếu tại trạm → Chấm điểm trải nghiệm (thang 1 - 5 sao)."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Phiếu ghi nhận số liệu thử nghiệm (Testing Data Sheet - Phụ lục 2) của mỗi trạm với đầy đủ con số thống kê (Tỷ lệ quét thành công %, Tỷ lệ hiểu bài %, Điểm âm thanh trung bình).")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act3 = doc.add_table(rows=4, cols=2)
    tbl_act3.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act3.autofit = False
    set_table_borders(tbl_act3, color="B0C4DE", sz="4")

    act3_steps = [
        ("Bước 1: Giao nhiệm vụ (3 phút)", "• Giáo viên phổ biến quy tắc Gallery Walk: Di chuyển trật tự, đeo tai nghe khi nghe audio để giữ yên tĩnh lớp học; phát Phiếu thu thập số liệu (Phụ lục 2)."),
        ("Bước 2: Thực hiện nhiệm vụ (15 phút)", "• Học sinh luân chuyển trải nghiệm ít nhất 3 - 4 trạm di tích.\n• Người trực trạm quan sát, ghi nhận số lượt quét thành công ngay lần đầu, thu phiếu trắc nghiệm của khách.\n• Giáo viên theo dõi, khích lệ các em tự tin giao tiếp tiếng Anh khi trải nghiệm."),
        ("Bước 3: Báo cáo, thảo luận (4 phút)", "• Các nhóm tập hợp lại số liệu. Đại diện 2 nhóm báo cáo nhanh: Tỷ lệ quét thành công đạt bao nhiêu %? Khách có nghe rõ lời không? Có gặp sự cố kỹ thuật nào không?"),
        ("Bước 4: Kết luận, nhận định (3 phút)", "• Giáo viên biểu dương các sản phẩm hoạt động tốt; hướng dẫn học sinh dùng chính các con số này để chuyển sang bước cải tiến kỹ thuật.")
    ]

    for idx, (s_t, s_d) in enumerate(act3_steps):
        row = tbl_act3.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ------------------ HOẠT ĐỘNG 4 ------------------
    h_act4 = doc.add_paragraph()
    h_act4.paragraph_format.space_before = Pt(12); h_act4.paragraph_format.space_after = Pt(4)
    r = h_act4.add_run("4. HOẠT ĐỘNG 4: ĐÁNH GIÁ, CẢI TIẾN VÀ MỞ RỘNG (Tiết 2 – 20 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act4.add_run("“PHÂN TÍCH SỐ LIỆU ĐO ĐẠC, CẢI TIẾN KỸ THUẬT VÀ XÂY DỰNG BẢN ĐỒ SỐ DI SẢN LỚP 8”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Học sinh biết dựa vào số liệu thực tế để tìm nguyên nhân và cách khắc phục lỗi kỹ thuật (Engineering Iteration); hoàn thành Phiếu nhật ký cải tiến; tải sản phẩm lên Padlet bản đồ số 'Saigon Junior Heritage Map' của lớp.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Nhóm phân tích kết quả thử nghiệm: So sánh tỷ lệ quét và điểm âm thanh với mục tiêu ban đầu; xác định lỗi (mã QR nhỏ, link Drive chưa mở quyền công khai, đọc hơi nhanh...).\n"
        "• Hoàn thành Phiếu nhật ký cải tiến (Phụ lục 3).\n"
        "• Khắc phục nhanh: Cấp quyền link Google Drive, phóng to mã QR trên Canva hoặc ghi chú điều chỉnh âm lượng.\n"
        "• Báo cáo phản ánh 30 giây: 'Our biggest problem was... and we improved it by...'.\n"
        "• Học sinh truy cập vào Website Cổng học liệu số, nhấn nút 'Học sinh nộp bài' để đưa mã QR, file audio và kịch bản song ngữ của nhóm lên web."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Phiếu nhật ký cải tiến (Phụ lục 3) đã hoàn thành; Sản phẩm được tối ưu hóa phiên bản 2.0; Trang giới thiệu di tích của nhóm được xuất bản chính thức trên Website Cổng học liệu số.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act4 = doc.add_table(rows=4, cols=2)
    tbl_act4.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act4.autofit = False
    set_table_borders(tbl_act4, color="B0C4DE", sz="4")

    act4_steps = [
        ("Bước 1: Giao nhiệm vụ (2 phút)", "• Giáo viên phát Phiếu nhật ký cải tiến (Phụ lục 3); yêu cầu các nhóm tìm ra ít nhất 01 lỗi kỹ thuật và 01 giải pháp khắc phục dựa trên số liệu thu được."),
        ("Bước 2: Thực hiện nhiệm vụ (8 phút)", "• Học sinh thảo luận nhóm, điền phiếu cải tiến.\n• Thao tác chỉnh sửa trực tiếp trên điện thoại/Canva (sửa quyền link Drive, đổi màu mã QR cho tương phản hơn)."),
        ("Bước 3: Báo cáo, thảo luận (5 phút)", "• 2 - 3 nhóm phát biểu ngắn gọn (45 giây/nhóm) về bài học kinh nghiệm và cách nhóm đã khắc phục sự cố kỹ thuật."),
        ("Bước 4: Kết luận, nhận định (5 phút)", "• Giáo viên tổng kết dự án: Nhận xét kiến thức tiếng Anh, khen ngợi tinh thần hợp tác và sự tiến bộ của học sinh lớp 8.\n• Hướng dẫn học sinh xuất bản sản phẩm lên Website Cổng học liệu số để lưu trữ lâu dài, chia sẻ cho phụ huynh, du khách và liên kết vào Kho học liệu số dùng chung của Sở GD&ĐT TP.HCM.")
    ]

    for idx, (s_t, s_d) in enumerate(act4_steps):
        row = tbl_act4.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ------------------ HỆ THỐNG PHỤ LỤC ------------------
    doc.add_page_break()

    h_pl = doc.add_paragraph()
    h_pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h_pl.add_run("HỆ THỐNG CÁC PHỤ LỤC VÀ TÀI LIỆU THAM KHẢO CHO LỚP 8\n"); r.font.bold = True; r.font.size = Pt(13.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_pl.add_run("(Học liệu số hỗ trợ, Phiếu ghi nhận số liệu và Bộ Rubrics đánh giá)"); r.font.italic = True; r.font.size = Pt(11)

    # Phụ lục 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run("PHỤ LỤC 1: KHUNG KỊCH BẢN THUYẾT MINH MẪU LỚP 8 (BILINGUAL SCRIPT SCAFFOLD)\n"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p.add_run("Ví dụ mẫu về Bưu điện Trung tâm Sài Gòn (Saigon Central Post Office) - Chuẩn lớp 8: 95 từ (khoảng 50 giây đọc):")
    r.font.italic = True

    tbl_scaffold = doc.add_table(rows=6, cols=3)
    tbl_scaffold.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_scaffold.autofit = False
    set_table_borders(tbl_scaffold, color="B0C4DE", sz="4")

    scaffold_data = [
        ("Phần kịch bản", "Nội dung tiếng Anh (Script Lớp 8)", "Ghi chú ngữ pháp & Kỹ thuật"),
        ("1. Hook\n(8 giây)", "“Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station?”", "Câu hỏi gây tò mò, từ vựng dễ, ngữ điệu thân thiện, vui vẻ."),
        ("2. Location\n(10 giây)", "“It is the Saigon Central Post Office, located in the heart of District 1, right next to Notre Dame Cathedral.”", "Cấu trúc: located in the heart of...; next to..."),
        ("3. History &\nArchitecture\n(17 giây)", "“This famous monument was built between 1886 and 1891 during the French period. Inside, you can see a beautiful arched ceiling and old wooden telephone booths.”", "Ngữ pháp lớp 8: was built between...; inside, you can see..."),
        ("4. Fun fact\n(10 giây)", "“Here is a fun fact: Many visitors still come here every day to write and send traditional postcards to their families around the world!”", "Chi tiết thú vị gần gũi, câu văn ngắn gọn."),
        ("5. Advice &\nClosing\n(10 giây)", "“You should take some nice photos here and remember to keep the place clean. Thank you and enjoy your trip in Saigon!”", "Động từ khuyết thiếu: You should...; lời chào kết thúc lịch sự.")
    ]

    for idx, (c0_t, c1_t, c2_t) in enumerate(scaffold_data):
        row = tbl_scaffold.rows[idx]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width = Inches(1.5); c1.width = Inches(3.4); c2.width = Inches(1.8)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60); set_cell_margins(c2, 60, 60, 60, 60)

        if idx == 0:
            for c in [c0, c1, c2]: set_cell_background(c, "003366")
            p = c0.paragraphs[0]; r = p.add_run(c0_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c1.paragraphs[0]; r = p.add_run(c1_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p = c2.paragraphs[0]; r = p.add_run(c2_t); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            set_cell_background(c0, "F4F7FA")
            p = c0.paragraphs[0]; r = p.add_run(c0_t); r.font.bold = True; r.font.size = Pt(11)
            p = c1.paragraphs[0]; r = p.add_run(c1_t); r.font.size = Pt(11)
            p = c2.paragraphs[0]; r = p.add_run(c2_t); r.font.size = Pt(10.5); r.font.italic = True

    # Phụ lục 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    r = p.add_run("PHỤ LỤC 2: PHIẾU GHI NHẬN SỐ LIỆU THỬ NGHIỆM LỚP 8 (TESTING DATA SHEET)\n"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p.add_run("Tên nhóm / Trạm di tích: ................................................................ Lớp: 8...... Ngày thử nghiệm: ..../..../2026")
    r.font.italic = True

    tbl_test = doc.add_table(rows=6, cols=5)
    tbl_test.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_test.autofit = False
    set_table_borders(tbl_test, color="B0C4DE", sz="4")

    test_headers = ["STT", "Chỉ số kỹ thuật đo đạc", "Cách tính công thức (M - Toán 8)", "Mục tiêu ban đầu", "Kết quả thực tế đo được"]
    test_rows = [
        ("1", "Tỷ lệ quét QR thành công ngay lần đầu", "(Số lượt quét thành công / Tổng số lượt thử) x 100%", "≥ 90%", "........ / ........ = ........%"),
        ("2", "Thời gian quét mã QR trung bình", "Bấm giờ điện thoại từ lúc giơ máy đến khi mở link", "≤ 3 giây", "........ giây"),
        ("3", "Tỷ lệ bạn nghe trả lời đúng câu hỏi", "(Số bạn trả lời đúng ≥ 2/3 câu / Tổng số bạn thử) x 100%", "≥ 80%", "........ / ........ = ........%"),
        ("4", "Điểm đánh giá Độ rõ âm thanh (S)", "Tính điểm trung bình cộng từ các bạn (thang 1 - 5 sao)", "≥ 4.0 / 5.0", "........ / 5.0 sao"),
        ("5", "Điểm đánh giá Bảng in A4 đẹp mắt (A)", "Tính điểm trung bình cộng tính thẩm mỹ (thang 1 - 5 sao)", "≥ 4.0 / 5.0", "........ / 5.0 sao")
    ]

    for c_i, h_txt in enumerate(test_headers):
        cell = tbl_test.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 50, 50, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(10)

    for r_i, r_data in enumerate(test_rows):
        row = tbl_test.rows[r_i + 1]
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            set_cell_margins(cell, 50, 50, 50, 50)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(10)

    # Phụ lục 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    r = p.add_run("PHỤ LỤC 3: PHIẾU NHẬT KÝ CẢI TIẾN KỸ THUẬT LỚP 8 (ENGINEERING IMPROVEMENT LOG)\n"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p.add_run("Minh chứng về quy trình cải tiến kỹ thuật dựa trên dữ liệu thực tế đo được:")
    r.font.italic = True

    tbl_imp = doc.add_table(rows=4, cols=4)
    tbl_imp.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_imp.autofit = False
    set_table_borders(tbl_imp, color="B0C4DE", sz="4")

    imp_headers = ["Vấn đề phát hiện", "Nguyên nhân kỹ thuật", "Cách nhóm đã khắc phục", "Kết quả sau khi sửa"]
    imp_samples = [
        ("Ví dụ: 2 bạn không quét được mã QR.", "Mã QR in nhỏ (2x2 cm), dán ở góc khuất và bị bóng đèn phản chiếu.", "Phóng to mã QR lên 4.5 x 4.5 cm và đặt ở trung tâm bảng A4.", "100% các bạn quét được ngay lập tức."),
        ("Ví dụ: Giọng đọc có tiếng quạt vù vù và thở mạnh vào mic.", "Thu âm gần quạt máy và để điện thoại quá sát miệng.", "Tắt quạt máy khi thu âm, giữ mic cách miệng 15 cm chếch 45 độ.", "Âm thanh trong trẻo, điểm âm thanh tăng lên 4.8/5 sao."),
        ("[Ghi nhận của nhóm em]:\n........................................................", "[Nguyên nhân]:\n........................................................", "[Giải pháp nhóm em]:\n........................................................", "[Kết quả đo lại]:\n........................................................")
    ]

    for c_i, h_txt in enumerate(imp_headers):
        cell = tbl_imp.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 50, 50, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(10)

    for r_i, r_data in enumerate(imp_samples):
        row = tbl_imp.rows[r_i + 1]
        for c_i, val in enumerate(r_data):
            cell = row.cells[c_i]
            set_cell_margins(cell, 50, 50, 50, 50)
            if r_i < 2: set_cell_background(cell, "F9FAFC")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9.5)
            if r_i < 2: r.font.italic = True

    # Phụ lục 4: Rubrics
    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("PHỤ LỤC 4: BỘ CÔNG CỤ ĐÁNH GIÁ DÀNH CHO LỚP 8 (RUBRICS & CHECKLISTS)\n"); r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Rubric 1
    p_r1 = doc.add_paragraph()
    r = p_r1.add_run("1. Rubric Kịch bản Tiếng Anh Lớp 8 (30%):\n"); r.font.bold = True; r.font.size = Pt(11)

    tbl_r1 = doc.add_table(rows=5, cols=4)
    tbl_r1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_r1.autofit = False
    set_table_borders(tbl_r1, color="B0C4DE", sz="4")

    r1_headers = ["Tiêu chí", "Mức Tốt (8 - 10đ)", "Mức Đạt (5 - 7đ)", "Cần cố gắng (< 5đ)"]
    r1_data = [
        ("Nội dung giới thiệu", "Thông tin di tích chính xác, có câu hỏi mở đầu hay, có chi tiết thú vị.", "Thông tin đúng nhưng câu mở đầu đơn giản, chưa có chi tiết hấp dẫn.", "Thông tin quá ít, thiếu chính xác hoặc sao chép trên mạng."),
        ("Ngữ pháp & Từ vựng Lớp 8", "Dùng đúng từ vựng di tích; dùng đúng thì quá khứ đơn, câu bị động và should.", "Có từ vựng nhưng còn 1 - 2 lỗi ngữ pháp thì quá khứ đơn hoặc should.", "Sai nhiều lỗi ngữ pháp cơ bản, từ vựng chưa phù hợp."),
        ("Cấu trúc 5 phần", "Đủ 5 phần rõ ràng, chuyển ý tự nhiên giữa các câu.", "Có đủ các phần nhưng các câu còn hơi rời rạc.", "Thiếu các phần quan trọng (như thiếu lời khuyên hoặc chào kết)."),
        ("Độ dài chuẩn lớp 8", "Vừa vặn chuẩn 80 - 110 từ (đọc trong 45 - 60 giây).", "Hơi ngắn (60 - 79 từ) hoặc hơi dài (111 - 130 từ).", "Quá ngắn (< 60 từ) hoặc quá dài (> 130 từ).")
    ]

    for c_i, h_txt in enumerate(r1_headers):
        cell = tbl_r1.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 45, 45, 45, 45)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    for r_i, r_d in enumerate(r1_data):
        row = tbl_r1.rows[r_i + 1]
        for c_i, val in enumerate(r_d):
            cell = row.cells[c_i]
            set_cell_margins(cell, 45, 45, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)

    # Rubric 2
    p_r2 = doc.add_paragraph()
    p_r2.paragraph_format.space_before = Pt(8)
    r = p_r2.add_run("2. Rubric Audio Guide & Phát âm Lớp 8 (30%):\n"); r.font.bold = True; r.font.size = Pt(11)

    tbl_r2 = doc.add_table(rows=4, cols=4)
    tbl_r2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_r2.autofit = False
    set_table_borders(tbl_r2, color="B0C4DE", sz="4")

    r2_headers = ["Tiêu chí", "Mức Tốt (8 - 10đ)", "Mức Đạt (5 - 7đ)", "Cần cố gắng (< 5đ)"]
    r2_data = [
        ("Phát âm & Ngữ điệu", "Phát âm rõ ràng, bật đúng âm cuối (/s/, /z/, /ed/), ngữ điệu tự nhiên, vui vẻ.", "Phát âm nhìn chung rõ; đôi chỗ còn quên âm đuôi hoặc giọng hơi đều.", "Phát âm sai nhiều từ cơ bản, đọc ngập ngừng."),
        ("Tốc độ & Thời lượng", "Thời lượng chuẩn 45 - 60 giây; tốc độ vừa phải, ngắt nghỉ đúng chỗ.", "Thời lượng 40 - 70 giây; nói hơi nhanh hoặc hơi chậm một chút.", "Nói quá nhanh hoặc quá chậm, thời lượng không đạt chuẩn."),
        ("Chất lượng âm thanh (S)", "Âm thanh trong trẻo, không có tiếng ồn quạt máy hay rè hơi thở; nhạc nền êm ái.", "Âm thanh nghe được nhưng còn lẫn chút tiếng ồn nhẹ xung quanh.", "Âm thanh bị rè, ồn, khó nghe rõ lời đọc.")
    ]

    for c_i, h_txt in enumerate(r2_headers):
        cell = tbl_r2.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 45, 45, 45, 45)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    for r_i, r_d in enumerate(r2_data):
        row = tbl_r2.rows[r_i + 1]
        for c_i, val in enumerate(r_d):
            cell = row.cells[c_i]
            set_cell_margins(cell, 45, 45, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)

    # Checklist 3 & 4
    p_c3 = doc.add_paragraph()
    p_c3.paragraph_format.space_before = Pt(8)
    r = p_c3.add_run("3. Bảng kiểm Kỹ thuật Bảng in A4 & Mã QR (20%):\n"); r.font.bold = True; r.font.size = Pt(11)
    p_c3.add_run(
        "• Mã QR quét thành công ngay trong 3 giây từ khoảng cách 20 - 30 cm [Đạt / Chưa đạt]\n"
        "• File Google Drive mở quyền xem công khai (không bắt đăng nhập) [Đạt / Chưa đạt]\n"
        "• Bảng A4 trang trí đẹp mắt, hình ảnh di tích rõ nét, có đoạn tóm tắt chữ lớn cho người nghe kém [Đạt / Chưa đạt]"
    )

    p_c4 = doc.add_paragraph()
    p_c4.paragraph_format.space_before = Pt(6)
    r = p_c4.add_run("4. Đánh giá Làm việc nhóm & Tinh thần cải tiến (20%):\n"); r.font.bold = True; r.font.size = Pt(11)
    p_c4.add_run(
        "• Tất cả thành viên trong nhóm đều tham gia tích cực và đoàn kết [Tốt / Đạt / Cần cố gắng]\n"
        "• Ghi nhận số liệu đo đạc thực tế trung thực trong buổi triển lãm [Tốt / Đạt / Cần cố gắng]\n"
        "• Hoàn thành phiếu cải tiến và biết cách sửa lỗi sản phẩm [Tốt / Đạt / Cần cố gắng]"
    )

    # Signature Block
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sign.autofit = False
    tbl_sign.rows[0].cells[0].width = Inches(3.2); tbl_sign.rows[0].cells[1].width = Inches(3.5)

    c0, c1 = tbl_sign.rows[0].cells[0], tbl_sign.rows[0].cells[1]
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DUYỆT CỦA TỔ CHUYÊN MÔN THCS\n"); r.font.bold = True
    r = p.add_run("(Ký và ghi rõ họ tên)"); r.font.italic = True

    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TP. Hồ Chí Minh, ngày       tháng 11 năm 2026\n"); r.font.italic = True
    r = p.add_run("GIÁO VIÊN SOẠN BÀI DẠY\n"); r.font.bold = True
    r = p.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"); r.font.italic = True
    r = p.add_run("[Họ và tên giáo viên]"); r.font.bold = True

    try:
        doc.save(filename)
        print(f"File {filename} created successfully.")
    except PermissionError:
        alt_name = "Tienganh_THCS_Ho-Va-Ten_STEM_P2_CapNhatWeb.docx"
        doc.save(alt_name)
        print(f"File {filename} is currently open in Word. Saved updated version as {alt_name} instead.")

if __name__ == "__main__":
    build_part_2_grade8()
