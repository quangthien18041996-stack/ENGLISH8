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

    # ------------------ TIÊU ĐỀ ĐẦU TRANG & HÀNH CHÍNH ------------------
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_top.autofit = False
    c_l, c_r = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_l.width = Inches(3.2); c_r.width = Inches(3.5)

    p_l = c_l.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_l.add_run("Trường: THCS Hòa Hiệp, TP. Hồ Chí Minh\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p_l.add_run("Tổ bộ môn: Tiếng Anh"); r.font.bold = True; r.font.size = Pt(11)

    p_r = c_r.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_r.add_run("Họ và tên giáo viên:\n"); r.font.bold = True; r.font.size = Pt(11)
    r = p_r.add_run("Trần Thị Thanh Nguyên"); r.font.size = Pt(11)

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
    r3 = p_box.add_run("Môn học: Tiếng Anh   |   Lớp dạy: 8A2 (Cấp THCS)   |   Thời gian thực hiện: 02 tiết (90 phút)\n")
    r3.font.bold = True; r3.font.size = Pt(11.5)
    r4 = p_box.add_run("(Tích hợp Unit 5: Our Customs and Traditions & Unit 11: Science and Technology – Tiếng Anh 8 Global Success)")
    r4.font.italic = True; r4.font.size = Pt(11)

    # ------------------ I. MỤC TIÊU BÀI HỌC ------------------
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14); h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("I. MỤC TIÊU BÀI HỌC"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Về kiến thức:\n"); r.font.bold = True
    p.add_run(
        "• Môn Tiếng Anh 8 (Global Success):\n"
        "  - Từ vựng: Vận dụng thành thạo từ vựng về di sản, danh lam thắng cảnh, văn hóa và du lịch TP.HCM: heritage, architecture, historical monument, post office, cathedral, market, wharf, attract, preserve, audio guide, tourist, colonial period...\n"
        "  - Ngữ pháp: Sử dụng chính xác thì quá khứ đơn (Past Simple), câu bị động (Passive Voice: was/were built, was established...), động từ khuyết thiếu đưa ra lời khuyên cho du khách (should, shouldn't), và cấu trúc miêu tả địa điểm (is located in, is famous for...).\n"
        "  - Kỹ năng viết kịch bản thuyết minh: Nắm vững cấu trúc 5 bước (Hook → Location → History/Architecture → Fun Fact → Advice/Closing) với dung lượng chuẩn 80 - 110 từ.\n"
        "• Tích hợp STEM / Dạy học số:\n"
        "  - Khoa học & Âm học (S): Hiểu về âm thanh, nguồn âm và tiêu âm; biết cách hạn chế tiếng ồn và tạp âm khi thu âm bằng điện thoại (giữ khoảng cách micrô 15 cm chếch 45 độ, thu âm trong phòng yên tĩnh < 45dB).\n"
        "  - Công nghệ & Tin học số (T): Biết tìm kiếm, đánh giá và chọn lọc hình ảnh số chất lượng cao; sử dụng Cổng học liệu số Di sản Sài Gòn (Web Portal); thực hiện tải ảnh, bài viết và file âm thanh lên hệ thống qua cổng nộp bài trực tuyến; thao tác quét mã QR và tương tác thả tim, bình luận thời gian thực trên nền tảng web di động.\n"
        "  - Kỹ thuật (E): Vận dụng quy trình thiết kế và hoàn thiện sản phẩm số: Nhận nhiệm vụ → Thu thập tư liệu & viết bài → Thu âm thử nghiệm → Giáo viên kiểm duyệt → Thuyết trình báo cáo tại lớp và tiếp thu phản hồi đồng đẳng.\n"
        "  - Nghệ thuật (A): Luyện phát âm chuẩn xác, diễn cảm, truyền cảm xúc tự hào về quê hương; chọn lọc hình ảnh trực quan, bố cục thẩm mỹ.\n"
        "  - Toán học (M): Kiểm soát thời lượng thuyết minh chuẩn 45 - 60 giây; đếm chính xác số lượng từ kịch bản (80 - 110 từ); thống kê số lượt tương tác thả tim và phản hồi bình luận thực tế tại lớp."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Về năng lực:\n"); r.font.bold = True
    p.add_run(
        "a) Năng lực đặc thù môn Tiếng Anh:\n"
        "  - Nghe: Nghe hiểu thông điệp và nội dung giới thiệu các di tích lịch sử của TP.HCM; nghe và đánh giá giọng phát âm, độ rõ ràng của nhóm bạn.\n"
        "  - Nói: Phát âm chuẩn các âm đuôi (/s/, /z/, /ed/), có ngữ điệu tự nhiên; tự tin đại diện nhóm thuyết trình giới thiệu di tích trước lớp ở Tiết 2.\n"
        "  - Đọc: Đọc hiểu và chọn lọc các tư liệu lịch sử, kiến trúc về địa danh được phân công.\n"
        "  - Viết: Viết bài giới thiệu tiếng Anh 80 - 110 từ súc tích, mạch lạc, đúng ngữ pháp và giàu sức gợi cảm.\n"
        "b) Năng lực chung:\n"
        "  - Tự chủ và tự học: Chủ động tìm tư liệu, luyện đọc kịch bản nhiều lần và tự giác hoàn thành sản phẩm thu âm tại nhà.\n"
        "  - Giao tiếp và hợp tác: Làm việc nhóm hiệu quả theo 5 vai trò chuyên trách, biết lắng nghe và hỗ trợ lẫn nhau trong quá trình thực hiện nhiệm vụ.\n"
        "  - Giải quyết vấn đề và sáng tạo: Sáng tạo câu mở đầu hấp dẫn (Hook), tìm chi tiết thú vị (Fun Fact) và ứng dụng công nghệ để tạo ra sản phẩm thuyết minh số hiện đại.\n"
        "c) Năng lực số: Thao tác thuần thục điện thoại thông minh để ghi âm, quét mã QR, truy cập Cổng học liệu số và tương tác bình luận học tập văn minh."
    )

    p = doc.add_paragraph()
    r = p.add_run("3. Về phẩm chất:\n"); r.font.bold = True
    p.add_run(
        "• Yêu nước & Tự hào dân tộc: Yêu quý các di tích lịch sử, văn hóa TP.HCM; có ý thức gìn giữ, phát huy giá trị di sản và mong muốn quảng bá nét đẹp quê hương đến bạn bè quốc tế.\n"
        "• Trách nhiệm: Nghiêm túc hoàn thành phần việc được phân công trong nhóm; tuân thủ quy định khi thu âm và nộp bài đúng thời hạn.\n"
        "• Chăm chỉ: Tỉ mỉ rèn luyện phát âm tiếng Anh, kiên nhẫn thu âm để đạt bản nghe trong trẻo và truyền cảm nhất."
    )

    # ------------------ II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU ------------------
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14); h2.paragraph_format.space_after = Pt(4)
    r = h2.add_run("II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p = doc.add_paragraph()
    r = p.add_run("1. Thiết bị và học liệu của Giáo viên:\n"); r.font.bold = True
    p.add_run(
        "• Thiết bị: Máy tính xách tay kết nối Internet, máy chiếu (hoặc màn hình tương tác thông minh), hệ thống loa lớp học, mạng Wi-Fi ổn định.\n"
        "• Hệ sinh thái Cổng học liệu số Di sản Sài Gòn (Web Portal):\n"
        "  - Địa chỉ Website chính thức: https://english8-two.vercel.app/ (Mã nguồn GitHub: https://github.com/quangthien18041996-stack/ENGLISH8).\n"
        "  - Tính năng Giáo viên kiểm duyệt (Teacher Verification): Giao diện bảo mật cho phép GV thẩm định, nhận xét, chỉnh sửa kịch bản tiếng Anh, ảnh và bản thu âm của học sinh nộp về; phê duyệt xuất bản chính thức lên trang web với huy hiệu xác nhận chất lượng 'Thông tin đã được phê duyệt ✓'.\n"
        "  - Tính năng Chế độ thuyết trình (Presentation Mode): Hỗ trợ trình chiếu toàn màn hình 4 trạm di tích song ngữ Anh - Việt trên máy chiếu lớp học; tích hợp Audio Player trực tiếp phát giọng đọc của học sinh; hiển thị mã QR Code phóng to (240x240px) để cả lớp dễ dàng quét mã tương tác.\n"
        "  - Tính năng Đồng bộ đám mây thời gian thực (Real-time Cloud Sync): Hệ thống tự động nhận các lượt bấm '❤️ Thả tim' và bình luận nhận xét (Peer Feedback) từ điện thoại của cả lớp, lập tức hiển thị hiệu ứng tim bay bùng nổ và thông báo bình luận trực tiếp trên màn hình máy chiếu lúc nhóm đang thuyết trình.\n"
        "  - Cổng tiếp nhận bài của học sinh (Student Submit): Hệ thống biểu mẫu cho phép các nhóm tải ảnh, bài viết tiếng Anh và file âm thanh lên web.\n"
        "• Học liệu bổ trợ đính kèm:\n"
        "  - Bộ 4 Phiếu học tập số hóa phân hóa cho 4 nhóm di tích (Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) trong tệp văn bản: Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx.\n"
        "  - Slide tương tác bài dạy, 02 file audio mẫu đối chiếu (Audio đạt chuẩn vs Audio lỗi tạp âm/phát âm) và Bộ tiêu chí chấm điểm (Rubric)."
    )

    p = doc.add_paragraph()
    r = p.add_run("2. Thiết bị và học liệu của Học sinh:\n"); r.font.bold = True
    p.add_run(
        "• Mỗi nhóm học tập (4 - 5 học sinh): 01 - 02 điện thoại thông minh (smartphone) có camera quét mã QR và kết nối Internet (Wi-Fi/4G); tai nghe cá nhân.\n"
        "• Phiếu học tập nhóm (in từ tệp Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx tương ứng với địa danh được phân công).\n"
        "• Tư liệu ảnh, bài viết và file âm thanh tiếng Anh đã được giáo viên phê duyệt kiểm duyệt, sẵn sàng cho buổi thuyết trình Tiết 2."
    )

    # ------------------ III. TIẾN TRÌNH DẠY HỌC ------------------
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14); h3.paragraph_format.space_after = Pt(4)
    r = h3.add_run("III. TIẾN TRÌNH DẠY HỌC CHI TIẾT"); r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # ================== HOẠT ĐỘNG 1 ==================
    h_act1 = doc.add_paragraph()
    h_act1.paragraph_format.space_before = Pt(12); h_act1.paragraph_format.space_after = Pt(4)
    r = h_act1.add_run("1. HOẠT ĐỘNG 1: XÁC ĐỊNH VẤN ĐỀ THỰC TIỄN / MỞ ĐẦU (Tiết 1 – 10 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act1.add_run("“RÀO CẢN NGÔN NGỮ CỦA DU KHÁCH VÀ DỰ ÁN ĐẠI SỨ VĂN HÓA NHÍ TP.HCM”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Tạo hứng thú học tập; khơi dậy niềm tự hào về di sản văn hóa TP.HCM; giúp học sinh nhận diện nhu cầu thuyết minh tiếng Anh của du khách quốc tế; phân chia 4 nhóm học tập và giao nhiệm vụ phụ trách 4 địa danh lịch sử trọng điểm của thành phố.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run("Học sinh quan sát hình ảnh/video ngắn về du khách nước ngoài tham quan các di tích TP.HCM; thảo luận câu hỏi định hướng: 'Làm thế nào để học sinh lớp 8 trở thành những Đại sứ văn hóa nhí (Junior Heritage Guides) giới thiệu nét đẹp quê hương cho bạn bè quốc tế?'; GV giới thiệu dự án số hóa và phân công 4 nhóm đảm nhận 4 địa danh.")

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("4 nhóm học sinh được thành lập, bầu nhóm trưởng và nhận nhiệm vụ phụ trách 1 trong 4 địa danh (Nhóm 1: Bưu điện Trung tâm Sài Gòn; Nhóm 2: Nhà thờ Đức Bà; Nhóm 3: Chợ Bến Thành; Nhóm 4: Bến Nhà Rồng).")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act1 = doc.add_table(rows=4, cols=2)
    tbl_act1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act1.autofit = False
    set_table_borders(tbl_act1, color="B0C4DE", sz="4")

    act1_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (3 phút)",
         "• Giáo viên chiếu video ngắn (20 giây) hoặc hình ảnh du khách nước ngoài đang tham quan di tích lịch sử tại trung tâm TP.HCM.\n"
         "• Giáo viên đặt câu hỏi gợi mở: 'Nếu bắt gặp một du khách quốc tế đang băn khoăn trước một di tích lịch sử của TP.HCM, em có thể làm gì để giúp họ hiểu được vẻ đẹp và ý nghĩa của nơi này chỉ trong 1 phút bằng tiếng Anh?'\n"
         "• Giới thiệu tên dự án học tập: 'Saigon Heritage QR Audio Guide – Voices of Our Community'.\n"
         "• Phân chia lớp thành 4 nhóm học tập, bốc thăm nhận 4 địa danh di tích lịch sử: Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng."),
        ("Bước 2: Thực hiện nhiệm vụ (3 phút)",
         "• Học sinh ngồi theo nhóm, thảo luận nhanh để bầu Nhóm trưởng (Team Leader) và Thư ký.\n"
         "• Các nhóm nhận diện địa danh của nhóm mình, thảo luận những hiểu biết ban đầu về địa danh được giao."),
        ("Bước 3: Báo cáo, thảo luận (2 phút)",
         "• Đại diện 4 nhóm đứng dậy xác nhận địa danh được phân công và nêu 1 câu cảm nghĩ ngắn gọn bằng tiếng Anh về địa danh đó (ví dụ: 'We are Team 1, we love Saigon Central Post Office because of its European architecture!')."),
        ("Bước 4: Kết luận, nhận định (2 phút)",
         "• Giáo viên chốt lại danh sách 4 nhóm và giới thiệu tổng quan tiến trình 2 tiết:\n"
         "  + Tiết 1: GV hướng dẫn tìm ảnh, viết bài giới thiệu và luyện phát âm, thu âm tiếng Anh.\n"
         "  + Ngoài giờ: Học sinh thu âm, gửi bài lên web để GV kiểm duyệt công bố.\n"
         "  + Tiết 2: Các nhóm lên thuyết trình trước lớp, GV chấm điểm, các bạn quét mã QR thả tim và gửi bình luận tương tác trực tiếp.")
    ]

    for idx, (s_t, s_d) in enumerate(act1_steps):
        row = tbl_act1.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ================== HOẠT ĐỘNG 2 ==================
    h_act2 = doc.add_paragraph()
    h_act2.paragraph_format.space_before = Pt(12); h_act2.paragraph_format.space_after = Pt(4)
    r = h_act2.add_run("2. HOẠT ĐỘNG 2: HƯỚNG DẪN TÌM ẢNH, VIẾT BÀI GIỚI THIỆU VÀ LUYỆN THU ÂM TIẾNG ANH (Tiết 1 – 35 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act2.add_run("“NGHIÊN CỨU KIẾN THỨC NỀN, XÂY DỰNG KỊCH BẢN VÀ CHUẨN HÓA KỸ THUẬT THU ÂM”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Học sinh nắm vững 3 kỹ năng cốt lõi: (1) Tiêu chí chọn ảnh tiêu biểu về di tích; (2) Cấu trúc 5 bước của bài giới thiệu tiếng Anh (80 - 110 từ) vận dụng ngữ pháp Unit 5 & 11; (3) Kỹ thuật phát âm chuẩn xác và kỹ thuật thu âm bằng điện thoại đảm bảo âm thanh trong trẻo, rõ ràng.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• GV hướng dẫn cách tìm ảnh: 1 ảnh toàn cảnh bao quát + 3 - 4 ảnh cận cảnh chi tiết kiến trúc đặc thù.\n"
        "• GV hướng dẫn cấu trúc bài viết tiếng Anh 5 bước: 1. Hook (Chào hỏi & câu hỏi gợi mở) → 2. Location (Vị trí) → 3. History/Architecture (Năm xây dựng, nét đặc sắc) → 4. Fun Fact (Điều thú vị) → 5. Advice/Closing (Lời khuyên & chào kết).\n"
        "• GV cung cấp ngân hàng từ vựng, ngữ pháp trọng tâm (past simple: was/were built, passive voice, should/shouldn't).\n"
        "• GV mở 2 file audio mẫu để đối chiếu: Audio đạt chuẩn (rõ lời, ngữ điệu, không tạp âm) vs Audio lỗi (tiếng thở mạnh vào mic, quạt gió, nuốt âm cuối /s/, /z/, /ed/, đọc quá nhanh).\n"
        "• Nhóm phân công 5 vai trò; thực hành viết nháp bài giới thiệu tiếng Anh trên Phiếu học tập số hóa."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Bản thảo bài giới thiệu tiếng Anh (80 - 110 từ) của 4 nhóm di tích được hoàn thiện trên Phiếu học tập; kế hoạch phân công thành viên tìm ảnh và thu âm ngoài giờ.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act2 = doc.add_table(rows=4, cols=2)
    tbl_act2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act2.autofit = False
    set_table_borders(tbl_act2, color="B0C4DE", sz="4")

    act2_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (5 phút)",
         "• Giáo viên phát Phiếu học tập số hóa chuyên biệt cho 4 nhóm (tệp Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx).\n"
         "• Hướng dẫn phân công 5 vai trò chuyên trách trong nhóm:\n"
         "  1. Trưởng nhóm (Team Leader): Quản lý tiến độ chung, đếm số từ kịch bản.\n"
         "  2. Người tìm ảnh (Photo Researcher): Tìm và chọn lọc 3 - 5 ảnh đẹp, sắc nét về di tích.\n"
         "  3. Người viết bài (Scriptwriter): Soạn thảo bài giới thiệu tiếng Anh 80 - 110 từ theo 5 bước.\n"
         "  4. Phát thanh viên (Voice Actor): Luyện phát âm chuẩn, phụ trách thu âm đoạn audio 45 - 60 giây.\n"
         "  5. Người thuyết trình (Presenter): Chuẩn bị bài nói thuyết trình trước lớp ở Tiết 2.\n"
         "• Mở 2 đoạn audio mẫu đối chiếu; hướng dẫn quy tắc thu âm: giữ mic cách miệng 15 cm chếch 45 độ, thu âm trong phòng kín có rèm tiêu âm, thời lượng chuẩn 45 - 60 giây."),
        ("Bước 2: Thực hiện nhiệm vụ (20 phút)",
         "• Học sinh làm việc theo nhóm:\n"
         "  + Thảo luận chọn lọc các đặc điểm nổi bật nhất của di tích nhóm phụ trách.\n"
         "  + Người viết bài cùng các bạn soạn thảo bài giới thiệu tiếng Anh theo cấu trúc 5 bước vào Phiếu học tập.\n"
         "  + Người tìm ảnh tìm kiếm các góc ảnh kiến trúc tiêu biểu trên mạng Internet hoặc ảnh tự chụp.\n"
         "  + Phát thanh viên cùng nhóm trưởng đọc thử kịch bản, đếm số từ (80 - 110 từ), tra cứu phiên âm IPA và gạch chân các âm đuôi (/s/, /z/, /ed/) cần bật rõ.\n"
         "• Giáo viên đi quanh các nhóm quan sát, hướng dẫn phát âm các từ vựng khó (architecture, magnificent, historical monument, bustling...) và chỉnh sửa câu cú."),
        ("Bước 3: Báo cáo, thảo luận (6 phút)",
         "• Giáo viên mời đại diện 2 nhóm đọc thử đoạn kịch bản nháp trước lớp.\n"
         "• Học sinh các nhóm khác lắng nghe, nhận xét về độ dài, ngữ điệu và phát âm.\n"
         "• Giáo viên khen ngợi các ý tưởng hay, chỉnh sửa lỗi dùng từ và lỗi ngữ pháp phổ biến."),
        ("Bước 4: Kết luận, nhận định (4 phút)",
         "• Giáo viên chốt lại cấu trúc bài viết và tiêu chuẩn kỹ thuật kịch bản cho 4 nhóm.\n"
         "• Hướng dẫn chi tiết nhiệm vụ ngoài giờ lên lớp:\n"
         "  + Hoàn thiện bài viết tiếng Anh và tập hợp 3 - 5 ảnh đẹp.\n"
         "  + Phát thanh viên tiến hành thu âm audio guide (45 - 60 giây).\n"
         "  + Nhóm truy cập https://english8-two.vercel.app/, nhấn 'Student Submit' để nộp bài lên web cho giáo viên kiểm duyệt.")
    ]

    for idx, (s_t, s_d) in enumerate(act2_steps):
        row = tbl_act2.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ================== HOẠT ĐỘNG NGOÀI LỚP ==================
    h_ext = doc.add_paragraph()
    h_ext.paragraph_format.space_before = Pt(12); h_ext.paragraph_format.space_after = Pt(4)
    r = h_ext.add_run("HOẠT ĐỘNG TRẢI NGHIỆM NGOÀI LỚP HỌC (GIỮA TIẾT 1 VÀ TIẾT 2)\n")
    r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x70, 0x30, 0xA0)
    r = h_act2.add_run("“HOÀN THIỆN ẢNH, THU ÂM TIẾNG ANH, NỘP BÀI VÀ GIÁO VIÊN KIỂM DUYỆT CÔNG BỐ”")
    r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    p_ext = doc.add_paragraph()
    p_ext.add_run(
        "Học sinh và Giáo viên phối hợp thực hiện trên nền tảng số trước khi bước vào Tiết 2:\n"
        "1. Nhiệm vụ của Học sinh:\n"
        "   - Tuyển chọn 3 - 5 bức ảnh chụp chất lượng cao, rõ nét về di tích được phân công (bao gồm toàn cảnh và chi tiết đặc sắc).\n"
        "   - Phát thanh viên luyện đọc kịch bản nhiều lần; tiến hành thu âm bằng điện thoại trong phòng yên tĩnh (thời lượng chuẩn 45 - 60 giây, giọng đọc truyền cảm, rõ âm đuôi).\n"
        "   - Truy cập Cổng học liệu số https://english8-two.vercel.app/, nhấn nút 'Student Submit / Học sinh nộp bài', điền tên nhóm, tải ảnh, dán bài viết tiếng Anh và đính kèm file ghi âm để gửi về cho giáo viên.\n"
        "2. Nhiệm vụ của Giáo viên (Teacher Verification):\n"
        "   - Giáo viên truy cập giao diện kiểm duyệt 'Teacher Verification' trên trang web.\n"
        "   - Đọc thẩm định bài viết tiếng Anh, kiểm tra ảnh chụp, nghe bản ghi âm của từng nhóm.\n"
        "   - Chỉnh sửa các lỗi nhỏ (nếu có), nhập nhận xét động viên và bấm 'Phê duyệt' để công bố sản phẩm của 4 nhóm lên Cổng học liệu số với huy hiệu chuẩn chất lượng 'Thông tin đã được phê duyệt ✓', sẵn sàng cho buổi thuyết trình Tiết 2."
    )

    # ================== HOẠT ĐỘNG 3 ==================
    h_act3 = doc.add_paragraph()
    h_act3.paragraph_format.space_before = Pt(12); h_act3.paragraph_format.space_after = Pt(4)
    r = h_act3.add_run("3. HOẠT ĐỘNG 3: BÁO CÁO THUYẾT TRÌNH TẠI LỚP, QUÉT MÃ QR TƯƠNG TÁC VÀ ĐÁNH GIÁ (Tiết 2 – 35 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act3.add_run("“THUYẾT TRÌNH DI TÍCH TRƯỚC LỚP, THẢ TIM VÀ BÌNH LUẬN THỜI GIAN THỰC QUA MÃ QR”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Các nhóm tự tin lên thuyết trình trước lớp về địa danh được giao; rèn luyện kỹ năng nói và diễn đạt tiếng Anh trước đám đông; GV đánh giá và chấm điểm theo Rubric; cả lớp dùng điện thoại quét mã QR trên màn hình để nghe lại audio, bấm thả tim và gửi bình luận trực tiếp; tạo không khí tương tác số sôi nổi, hào hứng.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• Lần lượt 4 nhóm (Nhóm 1 đến Nhóm 4) lên bục giảng thuyết trình về địa danh của mình (mỗi nhóm khoảng 6 - 7 phút).\n"
        "• GV bật Chế độ thuyết trình (Presentation Mode) trên máy chiếu lớn: hiển thị hình ảnh đẹp, bài viết song ngữ của nhóm và phát bản thu âm Audio Guide tiếng Anh do học sinh thu âm.\n"
        "• Mã QR của địa danh hiển thị nổi bật trên màn hình máy chiếu.\n"
        "• Học sinh các nhóm dưới lớp dùng điện thoại thông minh quét mã QR truy cập trang web, bấm '❤️ Thả tim' và gửi bình luận nhận xét (Peer Feedback) trực tiếp trên điện thoại (ví dụ: 'Great pronunciation!', 'Clear voice!', 'Beautiful pictures!').\n"
        "• Hệ thống Real-time Cloud Sync hiển thị hiệu ứng tim bay bùng nổ và các dòng bình luận ngay lập tức trên màn hình máy chiếu tạo sự hứng khởi đặc biệt.\n"
        "• GV nhận xét, đặt câu hỏi giao lưu và chấm điểm bài thuyết trình theo Rubric."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("4 bài thuyết trình hoàn chỉnh của 4 nhóm di tích; dữ liệu tương tác thực tế (số lượt thả tim và bình luận trên màn hình máy chiếu); bảng điểm đánh giá của giáo viên.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act3 = doc.add_table(rows=4, cols=2)
    tbl_act3.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act3.autofit = False
    set_table_borders(tbl_act3, color="B0C4DE", sz="4")

    act3_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (3 phút)",
         "• Giáo viên thông báo quy chế thuyết trình và cách thức tương tác của lớp:\n"
         "  + Mỗi nhóm có 6 - 7 phút để đại diện lên thuyết trình và phát đoạn audio guide của nhóm.\n"
         "  + Khi nhóm bạn thuyết trình, toàn bộ học sinh bên dưới dùng smartphone quét mã QR trên màn hình máy chiếu để nghe lại bài, bấm '❤️ Thả tim' cổ vũ và gõ lời bình luận nhận xét tích cực bằng tiếng Anh.\n"
         "• Giáo viên kích hoạt Chế độ thuyết trình (Presentation Mode) trên máy chiếu lớn tại địa chỉ https://english8-two.vercel.app/."),
        ("Bước 2: Thực hiện nhiệm vụ (24 phút)",
         "• Lần lượt 4 nhóm lên bục giảng thuyết trình (khoảng 6 phút/nhóm):\n"
         "  1. Nhóm 1: Bưu điện Trung tâm Sài Gòn (Saigon Central Post Office).\n"
         "  2. Nhóm 2: Nhà thờ Đức Bà (Notre Dame Cathedral).\n"
         "  3. Nhóm 3: Chợ Bến Thành (Ben Thanh Market).\n"
         "  4. Nhóm 4: Bến Nhà Rồng - Bảo tàng Hồ Chí Minh (Dragon Wharf).\n"
         "• Diễn biến tại mỗi lượt thuyết trình:\n"
         "  + Đại diện nhóm thuyết trình tự tin giới thiệu về địa danh bằng tiếng Anh; bật tệp audio guide để cả lớp cùng lắng nghe giọng đọc truyền cảm của phát thanh viên nhóm mình.\n"
         "  + Học sinh dưới lớp quét mã QR trên màn hình máy chiếu, bấm nút '❤️ Thả tim' liên tục và gửi nhận xét đồng đẳng (Peer Comments).\n"
         "  + Màn hình máy chiếu của lớp học tự động hiển thị hàng loạt trái tim bay bùng nổ cùng các dòng tin nhắn bình luận thời gian thực (Real-time Cloud Sync) khiến không khí lớp học vô cùng hào hứng và sôi động!"),
        ("Bước 3: Báo cáo, thảo luận (5 phút)",
         "• Sau mỗi bài thuyết trình, GV mời đại diện 1 nhóm khác nêu nhận xét trực tiếp: Ấn tượng nhất với câu nói nào trong bài? Giọng phát âm của bạn có rõ ràng không?\n"
         "• Nhóm thuyết trình trả lời câu hỏi giao lưu ngắn từ giáo viên hoặc các bạn."),
        ("Bước 4: Kết luận, nhận định (3 phút)",
         "• Giáo viên nhận xét, chấm điểm trực tiếp cho từng nhóm theo Phiếu đánh giá Rubric (về bài viết, giọng phát âm, phong thái thuyết trình và mức độ tương tác).\n"
         "• Tuyên dương các nhóm có sự chuẩn bị chu đáo, phát âm tiếng Anh tự nhiên và nhận được nhiều lượt tương tác thả tim nhất.")
    ]

    for idx, (s_t, s_d) in enumerate(act3_steps):
        row = tbl_act3.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ================== HOẠT ĐỘNG 4 ==================
    h_act4 = doc.add_paragraph()
    h_act4.paragraph_format.space_before = Pt(12); h_act4.paragraph_format.space_after = Pt(4)
    r = h_act4.add_run("4. HOẠT ĐỘNG 4: TỔNG KẾT, TUYÊN DƯƠNG VÀ LAN TỎA SẢN PHẨM (Tiết 2 – 10 phút)\n")
    r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_act4.add_run("“VINH DANH CÁC ĐẠI SỨ VĂN HÓA NHÍ VÀ KẾ HOẠCH LAN TỎA DI SẢN SÀI GÒN”")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    p = doc.add_paragraph()
    r = p.add_run("a) Mục tiêu: "); r.font.bold = True
    p.add_run("Tổng kết toàn bộ dự án dạy học STEM; biểu dương tinh thần làm việc nhóm và năng lực tiếng Anh, năng lực số của học sinh; trao giải/tuyên dương nhóm xuất sắc; định hướng lan tỏa sản phẩm ra cộng đồng.")

    p = doc.add_paragraph()
    r = p.add_run("b) Nội dung: "); r.font.bold = True
    p.add_run(
        "• GV tổng kết kết quả chấm điểm của 4 nhóm; vinh danh: Nhóm có bài thuyết trình xuất sắc nhất, Nhóm có giọng đọc truyền cảm nhất, Nhóm được yêu thích nhất (nhiều tim nhất).\n"
        "• Học sinh chia sẻ cảm xúc và điều học được sau dự án (tự tin hơn khi nói tiếng Anh, yêu quý hơn lịch sử TP.HCM, biết cách ứng dụng công nghệ ghi âm và mã QR).\n"
        "• GV hướng dẫn học sinh lan tỏa sản phẩm: Chia sẻ đường link Cổng học liệu số https://english8-two.vercel.app/ cho cha mẹ, người thân và du khách; đề xuất nhà trường dán mã QR tại Thư viện và Bảng tin tiếng Anh của trường."
    )

    p = doc.add_paragraph()
    r = p.add_run("c) Sản phẩm: "); r.font.bold = True
    p.add_run("Bảng tổng kết thi đua của 4 nhóm; kế hoạch lan tỏa mã QR giới thiệu di tích của học sinh lớp 8 ra cộng đồng.")

    p = doc.add_paragraph()
    r = p.add_run("d) Tổ chức thực hiện:\n"); r.font.bold = True

    tbl_act4 = doc.add_table(rows=4, cols=2)
    tbl_act4.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_act4.autofit = False
    set_table_borders(tbl_act4, color="B0C4DE", sz="4")

    act4_steps = [
        ("Bước 1: Chuyển giao nhiệm vụ (2 phút)",
         "• Giáo viên yêu cầu cả lớp nhìn lại toàn bộ hành trình dự án: Từ lúc nhận nhiệm vụ, tìm ảnh, viết kịch bản, luyện phát âm, thu âm, gửi bài duyệt đến khi tự tin thuyết trình trên bục giảng."),
        ("Bước 2: Thực hiện nhiệm vụ (4 phút)",
         "• Học sinh thảo luận nhanh trong nhóm, cử đại diện phát biểu cảm nghĩ:\n"
         "  + Điều gì khiến em cảm thấy tự hào nhất trong dự án này?\n"
         "  + Kỹ năng tiếng Anh và kỹ năng số nào của em đã tiến bộ rõ rệt nhất?"),
        ("Bước 3: Báo cáo, thảo luận (2 phút)",
         "• Đại diện 2 học sinh phát biểu cảm nhận trước lớp: Cảm thấy rất hào hứng khi được làm 'Đại sứ văn hóa nhí', thấy tự tin hơn khi nghe giọng đọc của chính mình được phát qua loa trên nền web chuyên nghiệp."),
        ("Bước 4: Kết luận, nhận định (2 phút)",
         "• Giáo viên tổng kết, công bố điểm và trao huy hiệu khen thưởng cho 4 nhóm:\n"
         "  + Giải 'Bài thuyết trình xuất sắc nhất' (Best Presentation).\n"
         "  + Giải 'Giọng đọc truyền cảm nhất' (Best Voice Actor).\n"
         "  + Giải 'Nhóm được yêu thích nhất' (Most Loved Team – nhiều lượt thả tim nhất).\n"
         "• Giáo viên khuyến khích học sinh chia sẻ sản phẩm số lên trang cá nhân hoặc gửi cho bạn bè quốc tế, lan tỏa hình ảnh đẹp của di sản Sài Gòn - TP.HCM.")
    ]

    for idx, (s_t, s_d) in enumerate(act4_steps):
        row = tbl_act4.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2); c1.width = Inches(4.5)
        set_cell_margins(c0, 60, 60, 60, 60); set_cell_margins(c1, 60, 60, 60, 60)
        set_cell_background(c0, "F4F7FA")
        p0 = c0.paragraphs[0]; r = p0.add_run(s_t); r.font.bold = True; r.font.size = Pt(11)
        p1 = c1.paragraphs[0]; r = p1.add_run(s_d); r.font.size = Pt(11)

    # ------------------ HỆ THỐNG PHỤ LỤC & RUBRICS ------------------
    doc.add_page_break()

    h_pl = doc.add_paragraph()
    h_pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h_pl.add_run("HỆ THỐNG CÁC PHỤ LỤC VÀ BỘ CÔNG CỤ ĐÁNH GIÁ (RUBRICS)\n"); r.font.bold = True; r.font.size = Pt(13.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_pl.add_run("(Kèm theo Kế hoạch bài dạy Chủ đề STEM Tiếng Anh 8)"); r.font.italic = True; r.font.size = Pt(11)

    p_note_ws = doc.add_paragraph()
    p_note_ws.paragraph_format.space_before = Pt(4); p_note_ws.paragraph_format.space_after = Pt(8)
    r = p_note_ws.add_run("Lưu ý về Học liệu số: "); r.font.bold = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    p_note_ws.add_run("Toàn bộ hệ thống 04 Phiếu học tập chi tiết dành cho 4 nhóm (Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) được biên soạn hoàn chỉnh trong tệp văn bản đính kèm: ")
    r = p_note_ws.add_run("Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx"); r.font.bold = True; r.font.underline = True
    p_note_ws.add_run(". Dưới đây là khung kịch bản minh họa và bộ công cụ đánh giá (Rubric) chính thức của giáo viên.")

    # Phụ lục 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run("PHỤ LỤC 1: KHUNG KỊCH BẢN THUYẾT MINH MẪU LỚP 8 (5-STEP BILINGUAL SCRIPT)\n"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p.add_run("Ví dụ minh họa về Bưu điện Trung tâm Sài Gòn (Saigon Central Post Office) - Dung lượng chuẩn: 95 từ (khoảng 50 giây đọc):")
    r.font.italic = True

    tbl_scaffold = doc.add_table(rows=6, cols=3)
    tbl_scaffold.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_scaffold.autofit = False
    set_table_borders(tbl_scaffold, color="B0C4DE", sz="4")

    scaffold_data = [
        ("Phần kịch bản", "Nội dung tiếng Anh (Script Lớp 8)", "Ghi chú ngữ pháp & Kỹ thuật"),
        ("1. Hook\n(8 giây)", "“Hello and welcome to Ho Chi Minh City! Do you know which building looks like a classic European train station?”", "Câu hỏi gợi mở gây tò mò, ngữ điệu vui vẻ, thân thiện."),
        ("2. Location\n(10 giây)", "“It is the Saigon Central Post Office, located in the heart of District 1, right next to Notre Dame Cathedral.”", "Cấu trúc chỉ vị trí: located in the heart of...; next to..."),
        ("3. History &\nArchitecture\n(17 giây)", "“This famous monument was built between 1886 and 1891 during the French period. Inside, you can see a beautiful arched ceiling and old wooden telephone booths.”", "Ngữ pháp lớp 8: was built between...; inside, you can see..."),
        ("4. Fun fact\n(10 giây)", "“Here is a fun fact: Many visitors still come here every day to write and send traditional postcards to their families around the world!”", "Chi tiết thú vị, câu văn ngắn gọn, dễ nhớ."),
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

    # Phụ lục 2: BỘ RUBRICS ĐÁNH GIÁ CHÍNH THỨC
    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("PHỤ LỤC 2: BỘ CÔNG CỤ ĐÁNH GIÁ BÀI HỌC STEM (RUBRICS)\n"); r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p.add_run("Tổng điểm: 100 điểm (Quy đổi về thang điểm 10 theo quy định)"); r.font.italic = True

    # Rubric 1
    p_r1 = doc.add_paragraph()
    p_r1.paragraph_format.space_before = Pt(8)
    r = p_r1.add_run("1. Tiêu chí Bài viết giới thiệu Tiếng Anh (30 điểm):\n"); r.font.bold = True; r.font.size = Pt(11)

    tbl_r1 = doc.add_table(rows=5, cols=4)
    tbl_r1.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_r1.autofit = False
    set_table_borders(tbl_r1, color="B0C4DE", sz="4")

    r1_headers = ["Tiêu chí", "Mức Tốt (8 - 10đ)", "Mức Đạt (5 - 7đ)", "Cần cố gắng (< 5đ)"]
    r1_data = [
        ("Nội dung giới thiệu di tích", "Thông tin chính xác, câu hỏi mở đầu hấp dẫn, có chi tiết thú vị lôi cuốn người nghe.", "Thông tin đúng nhưng câu mở đầu đơn giản, chi tiết thú vị chưa rõ nét.", "Thông tin sơ sài, thiếu chính xác hoặc sao chép nguyên văn trên mạng."),
        ("Ngữ pháp & Từ vựng Lớp 8", "Dùng đúng từ vựng di tích; vận dụng chuẩn thì quá khứ đơn, câu bị động và lời khuyên should.", "Có từ vựng nhưng còn 1 - 2 lỗi nhỏ về thì quá khứ đơn hoặc cấu trúc should.", "Sai nhiều lỗi ngữ pháp cơ bản, từ vựng chưa phù hợp chủ đề."),
        ("Cấu trúc kịch bản 5 phần", "Đủ 5 phần rõ ràng (Hook, Location, History, Fun fact, Advice), chuyển ý tự nhiên.", "Có đủ các phần nhưng liên kết giữa các câu còn hơi rời rạc.", "Thiếu các phần quan trọng (như thiếu lời khuyên hoặc chào kết)."),
        ("Dung lượng chuẩn lớp 8", "Chuẩn 80 - 110 từ (phù hợp thời lượng đọc 45 - 60 giây).", "Hơi ngắn (60 - 79 từ) hoặc hơi dài (111 - 130 từ).", "Quá ngắn (< 60 từ) hoặc quá dài (> 130 từ).")
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
    r = p_r2.add_run("2. Tiêu chí Tệp âm thanh Audio Guide (30 điểm):\n"); r.font.bold = True; r.font.size = Pt(11)

    tbl_r2 = doc.add_table(rows=4, cols=4)
    tbl_r2.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_r2.autofit = False
    set_table_borders(tbl_r2, color="B0C4DE", sz="4")

    r2_headers = ["Tiêu chí", "Mức Tốt (8 - 10đ)", "Mức Đạt (5 - 7đ)", "Cần cố gắng (< 5đ)"]
    r2_data = [
        ("Phát âm & Ngữ điệu", "Phát âm rõ ràng, bật đúng âm cuối (/s/, /z/, /ed/), ngữ điệu tự nhiên, truyền cảm xúc.", "Phát âm nhìn chung rõ; đôi chỗ còn quên âm đuôi hoặc giọng hơi đều.", "Phát âm sai nhiều từ cơ bản, đọc ngập ngừng, phát âm thiếu chuẩn xác."),
        ("Tốc độ & Thời lượng", "Thời lượng chuẩn 45 - 60 giây; tốc độ vừa phải, ngắt nghỉ đúng cụm nghĩa.", "Thời lượng 40 - 70 giây; nói hơi nhanh hoặc hơi chậm một chút.", "Nói quá nhanh hoặc quá chậm, thời lượng không đạt chuẩn."),
        ("Chất lượng âm thanh", "Âm thanh trong trẻo, không lẫn tiếng ồn quạt máy hay rè hơi thở; âm lượng vừa nghe.", "Âm thanh nghe được nhưng còn lẫn chút tiếng ồn nhẹ môi trường.", "Âm thanh bị rè, ồn, khó nghe rõ lời đọc.")
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

    # Rubric 3: Kỹ năng thuyết trình trước lớp (20 điểm)
    p_r3 = doc.add_paragraph()
    p_r3.paragraph_format.space_before = Pt(8)
    r = p_r3.add_run("3. Tiêu chí Kỹ năng Thuyết trình trước lớp (20 điểm):\n"); r.font.bold = True; r.font.size = Pt(11)

    tbl_r3 = doc.add_table(rows=3, cols=4)
    tbl_r3.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_r3.autofit = False
    set_table_borders(tbl_r3, color="B0C4DE", sz="4")

    r3_headers = ["Tiêu chí", "Mức Tốt (8 - 10đ)", "Mức Đạt (5 - 7đ)", "Cần cố gắng (< 5đ)"]
    r3_data = [
        ("Phong thái & Sự tự tin", "Tác phong chững chạc, tự tin, mắt tương tác với người nghe (eye contact), giọng nói to rõ.", "Tương đối tự tin; thỉnh thoảng còn nhìn vào tài liệu hoặc màn hình.", "Rụt rè, giọng nói quá nhỏ, chỉ đọc lại tài liệu."),
        ("Trình bày & Phối hợp", "Phối hợp nhịp nhàng giữa lời nói, hình ảnh trình chiếu và việc phát tệp audio; trả lời tốt câu hỏi giao lưu.", "Trình bày mạch lạc nhưng phối hợp chiếu và phát audio còn lúng túng đôi chỗ.", "Thiếu sự chuẩn bị, phối hợp rời rạc giữa các thành viên.")
    ]

    for c_i, h_txt in enumerate(r3_headers):
        cell = tbl_r3.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 45, 45, 45, 45)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    for r_i, r_d in enumerate(r3_data):
        row = tbl_r3.rows[r_i + 1]
        for c_i, val in enumerate(r_d):
            cell = row.cells[c_i]
            set_cell_margins(cell, 45, 45, 45, 45)
            if c_i == 0: set_cell_background(cell, "F4F7FA")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(9)

    # Rubric 4: Tương tác số & Làm việc nhóm (20 điểm)
    p_r4 = doc.add_paragraph()
    p_r4.paragraph_format.space_before = Pt(8)
    r = p_r4.add_run("4. Tiêu chí Tương tác số & Làm việc nhóm (20 điểm):\n"); r.font.bold = True; r.font.size = Pt(11)
    p_r4.add_run(
        "• Nộp bài đúng hạn lên web và được Giáo viên kiểm duyệt phê duyệt thành công [Tốt: 5đ / Đạt: 3đ / Chưa đạt: 1đ]\n"
        "• Phân công 5 vai trò rõ ràng, các thành viên tích cực hợp tác và đoàn kết [Tốt: 5đ / Đạt: 3đ / Chưa đạt: 1đ]\n"
        "• Tích cực quét mã QR, bấm thả tim và gửi bình luận văn minh cho các nhóm bạn [Tốt: 5đ / Đạt: 3đ / Chưa đạt: 1đ]\n"
        "• Thu hút được nhiều lượt thả tim và phản hồi tích cực từ cả lớp khi thuyết trình [Tốt: 5đ / Đạt: 3đ / Chưa đạt: 1đ]"
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
    r = p.add_run("Trần Thị Thanh Nguyên"); r.font.bold = True

    try:
        doc.save(filename)
        print(f"File {filename} created successfully.")
    except PermissionError:
        alt_name = filename.replace(".docx", "_CapNhatWeb.docx")
        doc.save(alt_name)
        print(f"File {filename} is currently open in Word. Saved updated version as {alt_name} instead.")

if __name__ == "__main__":
    build_part_2_grade8()
