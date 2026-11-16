import os
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

def build_part_2_grade8(filename=None):
    if filename is None:
        target_dir = os.path.join("01_Ho_So_Du_Thi_SGDDT", "Folder 2 - Ke hoach bai day")
        os.makedirs(target_dir, exist_ok=True)
        filename = os.path.join(target_dir, "Tienganh_THCS_Ho-Va-Ten_STEM_P2.docx")

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
        "  - Khoa học (S - KHTN 7-8): Hiểu về âm thanh, nguồn âm, tiếng ồn và tiêu âm; biết cách hạn chế tạp âm khi ghi âm bằng cách sử dụng vật liệu tiêu âm (rèm cửa, phòng kín < 45dB) và giữ đúng cự ly micrô (15cm, góc 45 độ).\n"
        "  - Công nghệ (T - Tin học 8): Khai thác Cổng học liệu số Di sản Sài Gòn (Web Portal), tương tác thời gian thực đám mây (Cloud Sync), tạo và quét mã QR, chia sẻ tệp âm thanh số, nộp bài trực tuyến.\n"
        "  - Kỹ thuật (E): Vận dụng quy trình thiết kế kỹ thuật 5 bước (Nêu vấn đề → Thiết kế kịch bản & bảng mẫu → Chế tạo mẫu thử → Kiểm thử thực tế Gallery Walk → Cải tiến).\n"
        "  - Nghệ thuật (A - Mỹ thuật 8): Bố cục trang trí bảng A4 sinh động, tương phản màu sắc giúp mã QR nổi bật; rèn luyện giọng đọc diễn cảm, tự tin.\n"
        "  - Toán học (M - Toán 8): Thu thập dữ liệu thực tế, tính tỷ lệ phần trăm (%) quét mã thành công (≥ 90%), tỷ lệ người nghe trả lời đúng nội dung (≥ 80%), tính điểm trung bình thang đo 1 - 5 sao và thống kê lượt tương tác thời gian thực."
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
        "  - Giao tiếp và hợp tác: Hợp tác tích cực trong nhóm 4 - 5 em với 5 vai trò rõ ràng, hỗ trợ nhau hoàn thành sản phẩm.\n"
        "  - Giải quyết vấn đề và sáng tạo: Sáng tạo giải pháp số thông minh để quảng bá di tích quê hương đến du khách quốc tế.\n"
        "c) Năng lực số: Thao tác thuần thục điện thoại thông minh, mã QR, Cổng học liệu số Web Portal và công cụ tương tác số an toàn, đúng quy định."
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
        "• Thiết bị: Máy tính kết nối Internet, máy chiếu hoặc màn hình tương tác, hệ thống loa lớp học, mạng Wi-Fi băng thông ổn định.\n"
        "• Hệ sinh thái Cổng học liệu số trực tuyến (Saigon Heritage Digital Portal):\n"
        "  - Địa chỉ Website chính thức: https://english8-two.vercel.app/ (Mã nguồn mở lưu trữ tại: https://github.com/quangthien18041996-stack/ENGLISH8).\n"
        "  - Tính năng Chế độ thuyết trình (Presentation Mode): Trình chiếu toàn màn hình 4 trạm di tích song ngữ Anh - Việt, tích hợp Audio Player trực tiếp và QR Code phóng to 240x240px phục vụ trình chiếu lớp học.\n"
        "  - Tính năng Giáo viên kiểm duyệt (Teacher Verification): Giao diện bảo mật cho phép GV thẩm định, đánh giá và cấp huy hiệu xác thực chất lượng 'Thông tin đã được phê duyệt ✓' trước khi công bố sản phẩm của học sinh.\n"
        "  - Hệ thống Đồng bộ tương tác đám mây thời gian thực (Real-time Cloud Sync qua Server-Sent Events): Cho phép học sinh dùng điện thoại thả tim và gửi nhận xét đồng đẳng (Peer Feedback) trực tiếp, hiển thị hiệu ứng tức thì trên màn hình máy chiếu.\n"
        "  - Cổng tiếp nhận sản phẩm số hóa (Student Submission Form): Nơi học sinh nộp ảnh khảo sát kiến trúc, kịch bản song ngữ 5 bước và tệp âm thanh số.\n"
        "• Học liệu chuyên biệt đính kèm:\n"
        "  - Bộ 4 Phiếu học tập phân hóa chuyên sâu cho 4 trạm di tích (Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) trong tệp văn bản: Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx.\n"
        "  - Slide tương tác bài dạy, 02 tệp audio mẫu đối chứng (Audio 1: đạt chuẩn vs Audio 2: lỗi tạp âm, nuốt âm cuối, đọc nhanh) và Bộ tiêu chí Rubric đánh giá (Phụ lục 1, 2, 3, 4)."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Thiết bị và học liệu của Học sinh:\n"); r.font.bold = True
    p.add_run(
        "• Mỗi nhóm học tập (4 - 5 học sinh): 01 - 02 điện thoại thông minh (smartphone) có camera quét mã QR, kết nối Wi-Fi/4G và tai nghe cá nhân (đảm bảo môi trường âm thanh yên tĩnh khi kiểm thử Gallery Walk).\n"
        "• Phiếu học tập chuyên biệt theo trạm di tích được phân công (In từ tệp Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx).\n"
        "• Dụng cụ học tập nhóm: Bảng giấy khổ A4/A3, bút dạ màu, thước kẻ, kéo, keo dán hai mặt để hoàn thiện bảng thuyết minh di tích."
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
    p.add_run("Tạo hứng thú, giúp học sinh lớp 8 nhận diện rào cản ngôn ngữ của du khách quốc tế tại các di tích lịch sử TP.HCM; xác định nhiệm vụ kỹ thuật: thiết kế Bảng thuyết minh A4 có mã QR Audio Guide tiếng Anh 45 - 60 giây và thống nhất 4 tiêu chí kỹ thuật (Design Specifications) của sản phẩm.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run("Học sinh xem đoạn clip tình huống ngắn về du khách nước ngoài lúng túng trước bảng chỉ dẫn tiếng Việt tại Bưu điện TP.HCM; thảo luận cặp đôi trả lời Driving Question: 'Làm thế nào để giúp du khách hiểu câu chuyện di tích chỉ trong 1 phút bằng điện thoại?'; làm rõ 4 tiêu chí kỹ thuật cần đạt.")

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run(
        "Bảng 4 Tiêu chí kỹ thuật sản phẩm (Design Specifications) được thống nhất và ghi nhận vào vở:\n"
        "  1. Kịch bản tiếng Anh: Chuẩn 80 - 110 từ, cấu trúc 5 bước, sử dụng thì quá khứ đơn, câu bị động và lời khuyên (should).\n"
        "  2. Tệp âm thanh Audio Guide: Thời lượng 45 - 60 giây, phát âm rõ ràng, có ngữ điệu, hạn chế tối đa tạp âm môi trường (< 45dB).\n"
        "  3. Mã QR số hóa: Quét nhạy trong 3 giây từ cự ly 20 - 30 cm; liên kết mở xem trực tiếp không cần đăng nhập.\n"
        "  4. Bảng thông tin in A4: Thiết kế khoa học, bố cục hài hòa, có ảnh khảo sát di tích, mã QR nổi bật và đoạn tóm tắt chữ lớn cho người khiếm thính/người nghe kém."
    )

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act1 = doc.add_table(rows=4, cols=2)
    tbl_act1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act1.autofit = False
    set_table_borders(tbl_act1, color="B0C4DE", sz="4")

    act1_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (3 phút)",
         "• Giáo viên trình chiếu đoạn video ngắn (20 giây) về tình huống du khách quốc tế bối rối khi tham quan di tích lịch sử TP.HCM do rào cản ngôn ngữ.\n"
         "• Giáo viên đặt câu hỏi gợi mở: 'Nếu em là một Hướng dẫn viên nhí (Junior Heritage Guide), làm thế nào để giúp du khách khám phá di tích chỉ trong 1 phút bằng chiếc điện thoại thông minh?'\n"
         "• Giới thiệu dự án STEM: 'Saigon Heritage QR Audio Guide – Voices of Our Community'."),
        ("Bước 2: Thực hiện nhiệm vụ (3 phút)",
         "• Học sinh suy nghĩ cá nhân, sau đó thảo luận cặp đôi (Think-Pair-Share) để đề xuất giải pháp khả thi.\n"
         "• Giáo viên định hướng giải pháp tích hợp công nghệ số: Bảng thông tin A4 có gắn mã QR liên kết file thuyết minh âm thanh tiếng Anh."),
        ("Bước 3: Báo cáo, thảo luận (2 phút)",
         "• Đại diện 2 cặp học sinh chia sẻ ý tưởng trước lớp: Bảng giấy truyền thống chứa quá nhiều chữ du khách ngại đọc; giải pháp âm thanh qua mã QR giúp du khách vừa ngắm cảnh vừa nghe thuyết minh sinh động."),
        ("Bước 4: Kết luận, nhận định (2 phút)",
         "• Giáo viên tổng hợp ý kiến, chốt lại bài toán kỹ thuật của dự án.\n"
         "• Trình chiếu và phân tích 4 Tiêu chí kỹ thuật bắt buộc của sản phẩm mẫu, giúp học sinh nắm rõ yêu cầu cần đạt.")
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
    p.add_run("Học sinh phân tích, đối chiếu được bản thu âm đạt chuẩn và bản thu âm lỗi (về âm học và ngữ âm); nắm vững cấu trúc kịch bản 5 bước; viết hoàn chỉnh kịch bản tiếng Anh 80 - 110 từ theo di tích được phân công và phác thảo Storyboard bảng thuyết minh A4 của nhóm.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Nghe đối chiếu 2 đoạn audio mẫu 45 giây của GV, phát hiện lỗi âm học (tiếng gió quạt, mic để quá gần miệng) và phát âm (nuốt âm đuôi /s/, /z/, /ed/, đọc thiếu ngữ điệu).\n"
        "• Ôn tập và áp dụng kiến thức ngôn ngữ lớp 8 (Unit 5 & 11): was/were built in..., is famous for..., You should visit / take photos...\n"
        "• Viết kịch bản nhóm theo cấu trúc 5 bước (Phụ lục 1 & Phiếu học tập số): 1. Hook (Chào hỏi & câu hỏi gợi mở 8 - 10s) → 2. Location (Vị trí 10s) → 3. History/Architecture (Năm xây dựng, nét đặc sắc 15 - 17s) → 4. Fun fact (Chi tiết thú vị 10s) → 5. Advice (Lời khuyên & lời chào 10s).\n"
        "• Biên tập chéo (Peer-editing) giữa các nhóm để đếm từ (80 - 110 từ) và sửa lỗi chính tả, ngữ pháp.\n"
        "• Vẽ phác thảo Storyboard bảng A4: xác định vị trí dán mã QR (tối thiểu 4x4 cm), vị trí ảnh khảo sát và khung tóm tắt chữ lớn."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Kịch bản thuyết minh tiếng Anh 80 - 110 từ đã hoàn thiện và Bản vẽ Storyboard bảng thuyết minh A4 trên giấy của 4 nhóm di tích.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act2 = doc.add_table(rows=4, cols=2)
    tbl_act2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act2.autofit = False
    set_table_borders(tbl_act2, color="B0C4DE", sz="4")

    act2_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (5 phút)",
         "• Giáo viên mở 2 đoạn audio mẫu qua hệ thống loa lớp học; hướng dẫn học sinh phân tích lỗi âm học và phát âm.\n"
         "• Hướng dẫn quy tắc thu âm âm học (Acoustic Rules): Giữ mic cách miệng 15 cm, nói chếch góc 45 độ để tránh luồng hơi thở trực diện, chọn không gian yên tĩnh có vật liệu tiêu âm.\n"
         "• Phát Bộ Phiếu học tập chuyên biệt cho 4 nhóm (tệp Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx): Nhóm 1 (Bưu điện), Nhóm 2 (Nhà thờ Đức Bà), Nhóm 3 (Chợ Bến Thành), Nhóm 4 (Bến Nhà Rồng).\n"
         "• Hướng dẫn phân công 5 vai trò trong nhóm: Trưởng nhóm, Biên kịch, Phát thanh viên, Kỹ thuật viên QR/Audio, Thuyết trình viên."),
        ("Bước 2: Thực hiện nhiệm vụ (20 phút)",
         "• Các nhóm làm việc: Phân tích 5 ảnh khảo sát kiến trúc trên phiếu học tập; viết kịch bản tiếng Anh 80 - 110 từ theo cấu trúc 5 bước.\n"
         "• Một học sinh phụ trách kỹ thuật phác thảo Storyboard bố cục bảng A4.\n"
         "• Trao đổi kịch bản kiểm tra chéo (Peer-edit) giữa các nhóm dựa trên checklist chuẩn từ vựng và ngữ pháp Unit 5/11.\n"
         "• Giáo viên đi quanh các nhóm cố vấn, chỉnh sửa lỗi phát âm từ khó (architecture, monument, magnificent, century, preservation...)."),
        ("Bước 3: Báo cáo, thảo luận (6 phút)",
         "• 01 đại diện nhóm đọc thử kịch bản tiếng Anh trước lớp với ngữ điệu diễn cảm, tự nhiên.\n"
         "• 01 đại diện nhóm trình bày bản vẽ Storyboard, giải thích vị trí đặt mã QR và cách phối màu làm nổi bật nội dung.\n"
         "• Giáo viên và các nhóm nhận xét, góp ý hoàn thiện."),
        ("Bước 4: Kết luận, nhận định (4 phút)",
         "• Giáo viên chốt lại khung kịch bản và tiêu chuẩn kỹ thuật cho 4 nhóm.\n"
         "• Hướng dẫn chi tiết nhiệm vụ chế tạo ngoài lớp học: Thu âm tại nhà, tải file lên Cổng học liệu số https://english8-two.vercel.app/, tạo mã QR và in màu bảng A4 chuẩn bị cho Tiết 2.")
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
    r = h_ext.add_run("CHẾ TẠO SẢN PHẨM MẪU: THU ÂM TIẾNG ANH, NỘP BÀI LÊN CỔNG SỐ VÀ IN BẢNG A4")
    r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    p_ext = doc.add_paragraph()
    p_ext.add_run(
        "Học sinh thực hiện tại nhà (1 - 2 ngày) dưới sự theo dõi và cố vấn trực tuyến của giáo viên:\n"
        "1. Thu âm Audio Guide tiếng Anh: Thực hiện trong phòng kín yên tĩnh (kéo rèm để tiêu âm, tạp âm nền < 45dB); mic cách miệng 15 cm chếch 45 độ; đọc kịch bản diễn cảm với thời lượng chuẩn 45 - 60 giây.\n"
        "2. Tải sản phẩm lên Cổng học liệu số (Web Submission): Học sinh truy cập https://english8-two.vercel.app/, nhấn nút 'Student Submit / Học sinh nộp bài' để tải lên 5 ảnh khảo sát di tích, kịch bản song ngữ và tệp âm thanh ghi âm.\n"
        "3. Giáo viên kiểm duyệt (Teacher Verification): Giáo viên truy cập cổng phê duyệt, kiểm tra nội dung học thuật, phát âm và âm thanh; phê duyệt cấp badge 'Thông tin đã được phê duyệt ✓' để sản phẩm hiển thị trên hệ thống số.\n"
        "4. Tạo mã QR & Thiết kế bảng A4: Tạo mã QR dẫn đến trang di tích trực tuyến của nhóm (hoặc tải trực tiếp từ web portal); hoàn thiện bảng thông tin khổ A4 (Canva hoặc trang trí thủ công) với mã QR kích thước tối thiểu 4 x 4 cm, in màu sẵn sàng cho buổi triển lãm Tiết 2."
    )

    # ------------------ HOẠT ĐỘNG 3 ------------------
    h_act3 = doc.add_paragraph()
    h_act3.paragraph_format.space_before = Pt(12); h_act3.paragraph_format.space_after = Pt(4)
    r = h_act3.add_run("3. HOẠT ĐỘNG 3: THỬ NGHIỆM SẢN PHẨM, BÁO CÁO VÀ THẢO LUẬN (Tiết 2 – 25 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act3.add_run("“TRIỂN LÃM GALLERY WALK – TRẢI NGHIỆM QUÉT MÃ VÀ ĐỒNG BỘ THỜI GIAN THỰC”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Tổ chức cho học sinh lớp 8 trải nghiệm đóng vai du khách quốc tế; rèn luyện kỹ năng nghe hiểu tiếng Anh thực tế; kiểm tra chất lượng kỹ thuật của bảng thuyết minh và mã QR; tương tác đồng bộ thời gian thực (Real-time Cloud Sync) và thu thập số liệu đo đạc thực tế (Testing Data Sheet).")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Bố trí không gian lớp học thành 4 trạm di tích lịch sử (Trạm 1: Bưu điện Trung tâm Sài Gòn; Trạm 2: Nhà thờ Đức Bà; Trạm 3: Chợ Bến Thành; Trạm 4: Bến Nhà Rồng).\n"
        "• Giáo viên kích hoạt 'Chế độ thuyết trình (Presentation Mode)' trên máy chiếu trung tâm: trình chiếu song ngữ Anh - Việt, tích hợp Audio Player và QR code phóng to 240x240px.\n"
        "• 02 thành viên mỗi nhóm (Trưởng nhóm & Thuyết trình viên) trực trạm hướng dẫn và ghi chép số liệu đo đạc; các thành viên còn lại đeo tai nghe luân chuyển trải nghiệm các trạm.\n"
        "• Học sinh dùng điện thoại thông minh quét mã QR → Nghe Audio Guide tiếng Anh → Trả lời câu hỏi trắc nghiệm nghe hiểu → Bấm '❤️ Thả tim' và gửi nhận xét đồng đẳng (Peer Feedback) trên điện thoại.\n"
        "• Dữ liệu tương tác được hệ thống Cloud Sync (ntfy SSE) truyền trực tiếp lên máy chiếu: các trái tim bay và thông báo chúc mừng lập tức xuất hiện trên màn hình lớp học."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Phiếu ghi nhận số liệu thử nghiệm (Testing Data Sheet - Phụ lục 2) của 4 trạm với đầy đủ con số thống kê thực tế (Tỷ lệ quét thành công %, Tỷ lệ hiểu bài %, Điểm đánh giá âm thanh, Tổng số lượt thả tim thời gian thực).")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act3 = doc.add_table(rows=4, cols=2)
    tbl_act3.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act3.autofit = False
    set_table_borders(tbl_act3, color="B0C4DE", sz="4")

    act3_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (3 phút)",
         "• Giáo viên phổ biến thể lệ triển lãm Gallery Walk: Di chuyển văn minh, đeo tai nghe cá nhân để giữ không gian học tập yên tĩnh.\n"
         "• Bật Chế độ thuyết trình (Presentation Mode) trên máy chiếu; hướng dẫn cách quét mã QR trên bảng in hoặc trên màn hình lớn để trải nghiệm tương tác trực tuyến.\n"
         "• Phát Phiếu ghi nhận số liệu thử nghiệm (Phụ lục 2) cho người trực trạm."),
        ("Bước 2: Thực hiện nhiệm vụ (15 phút)",
         "• Học sinh luân chuyển trải nghiệm qua 4 trạm di tích: Quét mã QR, nghe audio guide tiếng Anh, hoàn thành câu hỏi nghe hiểu trên điện thoại.\n"
         "• Học sinh bấm nút '❤️ Thả tim' và gửi nhận xét đánh giá đồng đẳng (Peer Review) trên giao diện web di động.\n"
         "• Màn hình máy chiếu của lớp học lập tức hiển thị hiệu ứng thả tim bay và thông báo phản hồi thời gian thực qua công nghệ Cloud Sync.\n"
         "• Người trực trạm quan sát, bấm giờ tốc độ quét QR và ghi lại con số thống kê thực tế vào Phiếu Phụ lục 2.\n"
         "• Giáo viên giám sát tiến trình, động viên học sinh tự tin giao tiếp và nhận xét bằng tiếng Anh."),
        ("Bước 3: Báo cáo, thảo luận (4 phút)",
         "• Các nhóm tập hợp lại số liệu đo đạc từ phiếu thử nghiệm. Đại diện 2 nhóm báo cáo nhanh số liệu:\n"
         "  + Tỷ lệ quét thành công ngay lần đầu đạt bao nhiêu %?\n"
         "  + Tốc độ quét trung bình mất mấy giây?\n"
         "  + Tỷ lệ các bạn nghe hiểu đúng câu hỏi trắc nghiệm?\n"
         "  + Nhóm nhận được bao nhiêu lượt thả tim và những phản hồi kỹ thuật nổi bật nào từ các bạn?"),
        ("Bước 4: Kết luận, nhận định (3 phút)",
         "• Giáo viên biểu dương các sản phẩm hoạt động ổn định và có lượt tương tác cao.\n"
         "• Nhấn mạnh ý nghĩa của các con số thực nghiệm: Dữ liệu đo đạc chính là cơ sở khoa học để các kỹ sư nhí tiến hành bước cải tiến sản phẩm.")
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
    p.add_run("Học sinh biết dựa vào số liệu thực nghiệm để tìm ra nguyên nhân và giải pháp kỹ thuật khắc phục lỗi (Engineering Iteration Process); hoàn thành Phiếu nhật ký cải tiến (Phụ lục 3); tối ưu hóa sản phẩm phiên bản 2.0 và hoàn thiện Cổng học liệu số Di sản Sài Gòn phục vụ cộng đồng.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Nhóm phân tích kết quả thử nghiệm: So sánh tỷ lệ quét thành công và điểm âm thanh thực tế với mục tiêu ban đầu; xác định các lỗi kỹ thuật (kích thước mã QR, độ tương phản màu, phân quyền liên kết, âm lượng hoặc tốc độ đọc).\n"
        "• Hoàn thành Phiếu nhật ký cải tiến kỹ thuật (Phụ lục 3).\n"
        "• Thực hiện khắc phục nhanh: Căn chỉnh mã QR, cập nhật kịch bản hoặc tinh chỉnh âm lượng trên Cổng học liệu số.\n"
        "• Thuyết trình ngắn 30 giây (30-second Engineering Pitch): 'Our biggest technical problem was... and we improved it by...'.\n"
        "• Giáo viên công bố bộ 4 di tích đã được phê duyệt kiểm định chính thức trên hệ thống web https://english8-two.vercel.app/ để lưu trữ và chia sẻ lâu dài."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Phiếu nhật ký cải tiến kỹ thuật (Phụ lục 3) đã hoàn thành; Sản phẩm Audio Guide & Bảng A4 được tối ưu hóa phiên bản 2.0; Hệ sinh thái Cổng học liệu số Di sản Sài Gòn của lớp được hoàn thiện và xuất bản chính thức.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act4 = doc.add_table(rows=4, cols=2)
    tbl_act4.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act4.autofit = False
    set_table_borders(tbl_act4, color="B0C4DE", sz="4")

    act4_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (2 phút)",
         "• Giáo viên phát Phiếu nhật ký cải tiến (Phụ lục 3); yêu cầu các nhóm dựa vào số liệu thực tế tại trạm để tìm ra ít nhất 01 lỗi kỹ thuật cần tối ưu và đề xuất giải pháp kỹ thuật cụ thể."),
        ("Bước 2: Thực hiện nhiệm vụ (8 phút)",
         "• Học sinh thảo luận nhóm, phân tích nguyên nhân (về âm học, quang học, mã QR hoặc thiết kế đồ họa) và điền phiếu cải tiến.\n"
         "• Thao tác điều chỉnh trực tiếp trên điện thoại/Canva/Cổng học liệu số (phóng to mã QR, tăng độ tương phản, điều chỉnh âm lượng audio)."),
        ("Bước 3: Báo cáo, thảo luận (5 phút)",
         "• Đại diện các nhóm báo cáo ngắn gọn (45 giây/nhóm) về bài học kinh nghiệm kỹ thuật và cách nhóm đã khắc phục sự cố để nâng cao trải nghiệm người dùng."),
        ("Bước 4: Kết luận, nhận định (5 phút)",
         "• Giáo viên tổng kết toàn bộ dự án STEM: Đánh giá cao sự tiến bộ về kỹ năng tiếng Anh (phát âm, kịch bản, tự tin giao tiếp), năng lực số và tư duy kỹ thuật của học sinh lớp 8.\n"
         "• Hướng dẫn học sinh lan tỏa sản phẩm: Chia sẻ đường link Cổng học liệu số https://english8-two.vercel.app/ cho cha mẹ, người thân và du khách; đề xuất nhà trường tích hợp mã QR tại Thư viện và Góc di sản trường học; sẵn sàng đóng góp vào Kho học liệu số của Sở GD&ĐT TP.HCM.")
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

    p_note_ws = doc.add_paragraph()
    p_note_ws.paragraph_format.space_before = Pt(4); p_note_ws.paragraph_format.space_after = Pt(8)
    r = p_note_ws.add_run("Lưu ý: "); r.font.bold = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p_note_ws.add_run("Hệ thống 04 Phiếu học tập chuyên biệt dành riêng cho 4 nhóm di tích (Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) được biên soạn hoàn chỉnh trong tệp văn bản đính kèm: ")
    r = p_note_ws.add_run("Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx"); r.font.bold = True; r.font.underline = True
    p_note_ws.add_run(". Dưới đây là khung kịch bản minh họa, phiếu đo đạc và bộ công cụ đánh giá chi tiết theo quy định.")

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

    tbl_test = doc.add_table(rows=7, cols=5)
    tbl_test.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_test.autofit = False
    set_table_borders(tbl_test, color="B0C4DE", sz="4")

    test_headers = ["STT", "Chỉ số kỹ thuật đo đạc", "Cách tính công thức (M - Toán 8)", "Mục tiêu ban đầu", "Kết quả thực tế đo được"]
    test_rows = [
        ("1", "Tỷ lệ quét QR thành công ngay lần đầu", "(Số lượt quét thành công / Tổng số lượt thử) x 100%", "≥ 90%", "........ / ........ = ........%"),
        ("2", "Thời gian quét mã QR trung bình", "Bấm giờ điện thoại từ lúc giơ máy đến khi mở link", "≤ 3 giây", "........ giây"),
        ("3", "Tỷ lệ bạn nghe trả lời đúng câu hỏi", "(Số bạn trả lời đúng ≥ 2/3 câu / Tổng số bạn thử) x 100%", "≥ 80%", "........ / ........ = ........%"),
        ("4", "Điểm đánh giá Độ rõ âm thanh (S)", "Tính điểm trung bình cộng từ các bạn (thang 1 - 5 sao)", "≥ 4.0 / 5.0", "........ / 5.0 sao"),
        ("5", "Điểm đánh giá Bảng in A4 đẹp mắt (A)", "Tính điểm trung bình cộng tính thẩm mỹ (thang 1 - 5 sao)", "≥ 4.0 / 5.0", "........ / 5.0 sao"),
        ("6", "Lượt tương tác Thả tim thời gian thực", "Số lượt bấm '❤️ Thả tim' hiển thị đồng bộ qua máy chiếu", "≥ 15 lượt", "........ lượt thả tim")
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
    r = p_c3.add_run("3. Bảng kiểm Kỹ thuật Bảng in A4, Mã QR & Web Portal (20%):\n"); r.font.bold = True; r.font.size = Pt(11)
    p_c3.add_run(
        "• Mã QR quét thành công ngay trong 3 giây từ cự ly 20 - 30 cm [Đạt / Chưa đạt]\n"
        "• Sản phẩm được tải lên Cổng học liệu số https://english8-two.vercel.app/ và được Giáo viên kiểm duyệt phê duyệt ✓ [Đạt / Chưa đạt]\n"
        "• Bảng in A4 thiết kế bắt mắt, hình ảnh khảo sát sắc nét, có đoạn tóm tắt chữ lớn cho người nghe kém [Đạt / Chưa đạt]\n"
        "• Tương tác thả tim và nhận xét đồng đẳng thời gian thực hiển thị ổn định trên màn hình lớp học [Đạt / Chưa đạt]"
    )

    p_c4 = doc.add_paragraph()
    p_c4.paragraph_format.space_before = Pt(6)
    r = p_c4.add_run("4. Đánh giá Làm việc nhóm & Tinh thần cải tiến (20%):\n"); r.font.bold = True; r.font.size = Pt(11)
    p_c4.add_run(
        "• Tất cả thành viên trong nhóm đều hoàn thành tốt vai trò được phân công (Leader, Scriptwriter, Voice Actor, Tech, Presenter) [Tốt / Đạt / Cần cố gắng]\n"
        "• Ghi nhận số liệu đo đạc thực tế trung thực trong buổi triển lãm Gallery Walk [Tốt / Đạt / Cần cố gắng]\n"
        "• Hoàn thành phiếu cải tiến (Phụ lục 3) và tối ưu hóa thành công phiên bản 2.0 [Tốt / Đạt / Cần cố gắng]"
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
        alt_name = filename.replace(".docx", "_CapNhatWeb.docx")
        doc.save(alt_name)
        print(f"File {filename} is currently open in Word. Saved updated version as {alt_name} instead.")

if __name__ == "__main__":
    build_part_2_grade8()
