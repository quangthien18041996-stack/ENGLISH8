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

def build_registration_docx():
    doc = docx.Document()
    setup_document_styles(doc)

    # ------------------ PHẦN ĐẦU TRANG & QUỐC HIỆU ------------------
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_top.autofit = False
    c_l, c_r = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_l.width = Inches(3.3); c_r.width = Inches(3.4)

    p_l = c_l.paragraphs[0]; p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_l.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO TP. HỒ CHÍ MINH\n"); r.font.bold = True; r.font.size = Pt(9.5)
    r = p_l.add_run("TRƯỜNG THCS HÒA HIỆP\n"); r.font.bold = True; r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = p_l.add_run("TỔ BỘ MÔN: TIẾNG ANH"); r.font.bold = True; r.font.size = Pt(9.5)

    p_r = c_r.paragraphs[0]; p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n"); r.font.bold = True; r.font.size = Pt(9.5)
    r = p_r.add_run("Độc lập - Tự do - Hạnh phúc\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p_r.add_run("-----------------------"); r.font.size = Pt(9.5)

    # Tiêu đề văn bản
    p_ref = doc.add_paragraph()
    p_ref.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ref.paragraph_format.space_before = Pt(8); p_ref.paragraph_format.space_after = Pt(2)
    r = p_ref.add_run("HỘI THI THIẾT KẾ CHỦ ĐỀ STEM/STEAM VÀ CHỦ ĐỀ DẠY HỌC SỐ NĂM HỌC 2026 - 2027\n")
    r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    r = p_ref.add_run("(Thực hiện theo Kế hoạch số 12070/KH-SGDĐT ngày 09/11/2026 của Sở GD&ĐT TP. Hồ Chí Minh)")
    r.font.italic = True; r.font.size = Pt(10)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8); p_title.paragraph_format.space_after = Pt(4)
    r = p_title.add_run("PHIẾU ĐĂNG KÝ CHỦ ĐỀ DẠY HỌC STEM VÀ BÀI GIẢNG SỐ\n")
    r.font.bold = True; r.font.size = Pt(14); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Callout Box Tên chủ đề
    tbl_name = doc.add_table(rows=1, cols=1)
    tbl_name.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_name.rows[0].cells[0].width = Inches(6.7)
    set_callout_box(tbl_name, border_color="003366", bg_color="EBF3FA")
    p_box = tbl_name.rows[0].cells[0].paragraphs[0]; p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_box.add_run("TÊN CHỦ ĐỀ: "); r1.font.bold = True; r1.font.size = Pt(12); r1.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r2 = p_box.add_run("SAIGON HERITAGE QR AUDIO GUIDE – VOICES OF OUR COMMUNITY\n")
    r2.font.bold = True; r2.font.size = Pt(12.5); r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r3 = p_box.add_run("(Thiết kế bảng thuyết minh thông minh tích hợp mã QR hướng dẫn âm thanh tiếng Anh giới thiệu di tích lịch sử - văn hóa TP.HCM dành cho học sinh Lớp 8A2)\n")
    r3.font.italic = True; r3.font.size = Pt(10.5)
    r4 = p_box.add_run("Môn: Tiếng Anh 8 (Global Success)  |  Giáo viên: TRẦN THỊ THANH NGUYÊN  |  Trường THCS Hòa Hiệp")
    r4.font.bold = True; r4.font.size = Pt(10.5); r4.font.color.rgb = RGBColor(0x00, 0x66, 0x00)

    # ------------------ PHẦN A: BẢNG TỔNG HỢP THEO MẪU ------------------
    h_a = doc.add_paragraph()
    h_a.paragraph_format.space_before = Pt(12); h_a.paragraph_format.space_after = Pt(4)
    r = h_a.add_run("PHẦN A: BẢNG TỔNG HỢP THÔNG TIN ĐĂNG KÝ CHỦ ĐỀ DỰ THI\n"); r.font.bold = True; r.font.size = Pt(11.5); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_a.add_run("(Điền đầy đủ và chuẩn xác theo 10 cột trong biểu mẫu đăng ký của Hội thi)"); r.font.italic = True; r.font.size = Pt(10)

    tbl_reg = doc.add_table(rows=11, cols=3)
    tbl_reg.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_reg.autofit = False
    set_table_borders(tbl_reg, color="B0C4DE", sz="4")

    tbl_headers = ["STT / Cột thông tin", "Tiêu đề theo biểu mẫu", "Nội dung đăng ký chi tiết của Giáo viên"]
    for c_i, h_txt in enumerate(tbl_headers):
        cell = tbl_reg.rows[0].cells[c_i]
        set_cell_background(cell, "003366")
        set_cell_margins(cell, 45, 45, 50, 50)
        p = cell.paragraphs[0]; r = p.add_run(h_txt); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(9.5)

    tbl_reg.rows[0].cells[0].width = Inches(1.5)
    tbl_reg.rows[0].cells[1].width = Inches(1.8)
    tbl_reg.rows[0].cells[2].width = Inches(3.4)

    reg_fields = [
        ("Cột 1", "Tên chủ đề dự kiến", 
         "SAIGON HERITAGE QR AUDIO GUIDE – VOICES OF OUR COMMUNITY\n"
         "(Thiết kế bảng thuyết minh thông minh tích hợp mã QR hướng dẫn âm thanh tiếng Anh giới thiệu di tích lịch sử - văn hóa TP.HCM dành cho học sinh Lớp 8A2)."),
        
        ("Cột 2", "Vấn đề thực tiễn / chủ đề hướng tới", 
         "• Du khách quốc tế tham quan các di tích TP.HCM thường gặp rào cản ngôn ngữ, bảng thông tin truyền thống chỉ có tiếng Việt và câu chữ quá dài, thiếu trải nghiệm âm thanh sống động.\n"
         "• Bài toán thực tiễn: Giúp học sinh Lớp 8A2 hóa thân thành 'Đại sứ văn hóa nhí' (Junior Heritage Guides), tự sáng tạo bản thuyết minh âm thanh tiếng Anh 45 - 60 giây và mã QR, giúp du khách nghe hiểu câu chuyện di tích chỉ trong 1 phút bằng điện thoại thông minh."),
        
        ("Cột 3", "Định hướng thiết kế bài dạy", 
         "• Thiết kế theo định hướng giáo dục STEM gắn với chuyển đổi số (chuẩn Công văn 5512/BGDĐT-GDTrH và Kế hoạch 12070/KH-SGDĐT của Sở GD&ĐT TP.HCM).\n"
         "• Tích hợp liên môn: Tiếng Anh 8 Global Success (Unit 5 & Unit 11) kết hợp Khoa học (S - kiểm soát âm thanh < 45dB), Công nghệ (T - Web Portal, QR Code), Kỹ thuật (E - quy trình thiết kế kỹ thuật) và Toán học (M - kiểm soát 80-110 từ, thời lượng 45-60s, thống kê tương tác).\n"
         "• Tiến trình 2 tiết kết hợp ngoài lớp: Tiết 1 hướng dẫn tìm ảnh, viết bài và thu âm → Ngoài giờ HS thu âm, nộp bài, GV kiểm duyệt → Tiết 2 các nhóm thuyết trình trước lớp (Presentation Mode), GV chấm điểm theo Rubric, cả lớp quét mã QR nghe lại, thả tim và bình luận nhận xét trực tiếp."),
        
        ("Cột 4", "Học liệu số dự kiến sử dụng", 
         "• Website Cổng học liệu số Di sản Sài Gòn trực tuyến: https://english8-two.vercel.app/ (Mã nguồn GitHub: quangthien18041996-stack/ENGLISH8).\n"
         "• Chế độ thuyết trình (Presentation Mode) toàn màn hình trên máy chiếu, tích hợp Audio Player và QR code phóng to 240x240px.\n"
         "• Tính năng Giáo viên kiểm duyệt (Teacher Verification) bảo mật, cấp badge chuẩn chất lượng 'Thông tin đã được phê duyệt ✓'.\n"
         "• Hệ thống Đồng bộ đám mây thời gian thực (Real-time Cloud Sync): hiển thị tim bay và bình luận trực tiếp từ điện thoại HS lên màn hình máy chiếu lúc thuyết trình.\n"
         "• Bộ 4 Phiếu học tập số hóa chuyên sâu cho 4 trạm di tích (Bưu điện, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng)."),
        
        ("Cột 5", "Khối lớp dự kiến", 
         "Khối 8 (Lớp thực hiện: 8A2) – Cấp Trung học cơ sở (Trường THCS Hòa Hiệp)."),
        
        ("Cột 6", "Thời lượng (số tiết)", 
         "02 tiết dạy học chính khóa trên lớp (90 phút) + Hoạt động trải nghiệm thu âm, nộp bài ngoài giờ lên lớp (1 - 2 ngày giữa hai tiết)."),
        
        ("Cột 7", "Không gian thực hiện dự kiến", 
         "• Trong lớp: Phòng học bộ môn Tiếng Anh / Phòng học thông minh Lớp 8A2 (máy tính, máy chiếu/SmartTV, loa âm thanh, Wi-Fi).\n"
         "• Ngoài lớp: Góc học tập yên tĩnh tại nhà của học sinh (đảm bảo độ ồn < 45dB để thu âm chất lượng cao).\n"
         "• Không gian số: Nền tảng Web Portal trực tuyến 24/7 (https://english8-two.vercel.app/) phục vụ lưu trữ, chia sẻ cộng đồng và tích hợp kho học liệu Sở GD&ĐT TP.HCM."),
        
        ("Cột 8", "Thâm niên công tác (năm)", 
         "06 năm (Giảng dạy môn Tiếng Anh cấp THCS)."),
        
        ("Cột 9", "Số điện thoại liên hệ", 
         "0777.693.878 (Email: nguyenttt@brt.edu.vn)."),
        
        ("Cột 10", "Ghi chú / đề xuất của tổ trưởng", 
         "“Chủ đề có ý tưởng sáng tạo, thiết thực, bám sát Kế hoạch 12070 của Sở GD&ĐT TP.HCM. Kế hoạch bài dạy rõ ràng theo Công văn 5512, vừa sức học sinh lớp 8; hệ thống học liệu số chạy mượt mà, tính tương tác sư phạm cao, sản phẩm ứng dụng tốt vào thực tiễn địa phương. Tổ chuyên môn Tiếng Anh nhất trí duyệt và đề xuất dự thi cấp Thành phố.”")
    ]

    for r_i, (c0_t, c1_t, c2_t) in enumerate(reg_fields):
        row = tbl_reg.rows[r_i + 1]
        row.cells[0].width = Inches(1.5)
        row.cells[1].width = Inches(1.8)
        row.cells[2].width = Inches(3.4)
        for c_i, val in enumerate([c0_t, c1_t, c2_t]):
            cell = row.cells[c_i]
            set_cell_margins(cell, 35, 35, 45, 45)
            if c_i == 0: 
                set_cell_background(cell, "F4F7FA")
            elif c_i == 1:
                set_cell_background(cell, "FAFBFD")
            p = cell.paragraphs[0]; r = p.add_run(val); r.font.size = Pt(8.5)
            if c_i < 2: r.font.bold = True

    # ------------------ PHẦN B: BẢN GIẢI TRÌNH CHI TIẾT ------------------
    doc.add_page_break()
    h_b = doc.add_paragraph()
    h_b.paragraph_format.space_before = Pt(8); h_b.paragraph_format.space_after = Pt(4)
    r = h_b.add_run("PHẦN B: BẢN THUYẾT MINH CHI TIẾT TỪNG NỘI DUNG ĐĂNG KÝ\n"); r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    r = h_b.add_run("(Tư liệu phục vụ giải trình sư phạm trước Hội đồng chuyên môn và Ban Giám hiệu)"); r.font.italic = True; r.font.size = Pt(10)

    # 1. Thông tin giáo viên
    p = doc.add_paragraph()
    r = p.add_run("1. Thông tin Giáo viên đăng ký:\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Họ và tên: TRẦN THỊ THANH NGUYÊN   |   Năm sinh: 1989   |   Giới tính: Nữ\n"
        "• Đơn vị công tác: Trường Trung học Cơ sở Hòa Hiệp, TP. Hồ Chí Minh\n"
        "• Chức vụ / Bộ môn: Giáo viên môn Tiếng Anh (Tổ Tiếng Anh THCS)\n"
        "• Thâm niên giảng dạy: 06 năm công tác liên tục trong ngành Giáo dục\n"
        "• Số điện thoại: 0777.693.878   |   Email: nguyenttt@brt.edu.vn"
    )

    # 2. Tên chủ đề
    p = doc.add_paragraph()
    r = p.add_run("2. Tên chủ đề dạy học số & STEM dự kiến:\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Tên tiếng Anh: SAIGON HERITAGE QR AUDIO GUIDE – VOICES OF OUR COMMUNITY\n"
        "• Tên tiếng Việt: Xây dựng Bảng thuyết minh số tích hợp mã QR hướng dẫn âm thanh tiếng Anh giới thiệu di tích lịch sử - văn hóa TP.HCM dành cho học sinh Lớp 8A2.\n"
        "• Môn học tích hợp chính: Tiếng Anh 8 (Bộ sách Global Success - GDPT 2018)."
    )

    # 3. Vấn đề thực tiễn
    p = doc.add_paragraph()
    r = p.add_run("3. Vấn đề thực tiễn và tính cấp thiết của đề tài:\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Xuất phát điểm: TP.HCM đón hàng triệu lượt khách quốc tế mỗi năm, trong đó các di tích như Bưu điện Trung tâm Sài Gòn, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng là những điểm đến tiêu biểu. Tuy nhiên, du khách tự túc thường gặp trở ngại vì các bảng thông tin truyền thống chỉ có chữ tiếng Việt, nội dung quá dài dòng, thiếu thuyết minh âm thanh tiếng Anh súc tích.\n"
        "• Câu hỏi định hướng (Driving Question): “Làm thế nào để học sinh Lớp 8A2 có thể trở thành những Đại sứ văn hóa nhí (Junior Heritage Guides), giúp du khách quốc tế hiểu được câu chuyện đặc sắc của di tích TP.HCM chỉ trong 60 giây bằng tiếng Anh qua một thao tác quét mã QR tiện lợi ngay trên điện thoại?”\n"
        "• Ý nghĩa giáo dục: Vừa giúp học sinh phát triển năng lực giao tiếp tiếng Anh thực tế (vượt qua tâm lý ngại nói), vừa giáo dục lòng tự hào dân tộc, tình yêu quê hương và ý thức gìn giữ di sản địa phương."
    )

    # 4. Định hướng thiết kế bài dạy
    p = doc.add_paragraph()
    r = p.add_run("4. Định hướng thiết kế bài dạy & Cấu trúc sư phạm:\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Thiết kế bám sát Công văn 5512/BGDĐT-GDTrH: Mỗi hoạt động dạy học đều tuân thủ chặt chẽ 4 bước: (1) Chuyển giao nhiệm vụ → (2) Thực hiện nhiệm vụ → (3) Báo cáo, thảo luận → (4) Kết luận, nhận định.\n"
        "• Cấu trúc 2 tiết học trên lớp (90 phút) + Hoạt động ngoài lớp học:\n"
        "  - Tiết 1 (45 phút - Tại lớp): Hoạt động 1 (Khởi động, giao nhiệm vụ 4 nhóm phụ trách 4 di tích) + Hoạt động 2 (Hướng dẫn tìm ảnh kiến trúc, cấu trúc viết bài giới thiệu tiếng Anh 5 bước: Hook → Location → History/Architecture → Fun Fact → Advice: 80 - 110 từ, và kỹ thuật phát âm, thu âm trên điện thoại).\n"
        "  - Hoạt động ngoài lớp (1 - 2 ngày giữa hai tiết): Học sinh thu âm đoạn audio tiếng Anh 45 - 60 giây (< 45dB), truy cập web portal nhấn 'Student Submit' để gửi ảnh, bài viết và file ghi âm; Giáo viên sử dụng giao diện 'Teacher Verification' kiểm tra, duyệt và công bố chính thức sản phẩm.\n"
        "  - Tiết 2 (45 phút - Tại lớp): Hoạt động 3 (Lần lượt 4 nhóm lên bục giảng thuyết trình trước lớp bằng Chế độ Presentation Mode, GV chấm điểm theo Rubric, cả lớp dùng điện thoại quét mã QR nghe lại, bấm '❤️ Thả tim' và gửi bình luận nhận xét trực tiếp hiển thị tức thì trên máy chiếu) + Hoạt động 4 (Tổng kết, biểu dương nhóm thuyết trình xuất sắc và nhóm nhận được nhiều tim nhất, dặn dò kế hoạch lan tỏa di sản)."
    )

    # 5. Học liệu số dự kiến sử dụng
    p = doc.add_paragraph()
    r = p.add_run("5. Hệ thống học liệu số dự kiến sử dụng (Công nghệ số hóa):\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Cổng học liệu số Di sản Sài Gòn (Web Portal): https://english8-two.vercel.app/ hoạt động 24/7 trên nền tảng đám mây, tích hợp đầy đủ tính năng số hiện đại:\n"
        "  + Chế độ thuyết trình (Presentation Mode): Toàn màn hình máy chiếu, song ngữ Anh - Việt, tích hợp Audio Player trực tiếp và QR Code phóng to (240x240px).\n"
        "  + Hệ thống Giáo viên kiểm duyệt (Teacher Verification): Mã bảo mật xác thực nội dung, cấp badge 'Thông tin đã được phê duyệt ✓' đảm bảo chuẩn mực sư phạm trước khi công bố.\n"
        "  + Hệ thống Đồng bộ đám mây thời gian thực (Real-time Cloud Sync qua Server-Sent Events): Cho phép cả lớp tương tác thả tim và gửi nhận xét đồng đẳng tức thì lên máy chiếu.\n"
        "  + Cổng tiếp nhận sản phẩm số hóa (Student Submission Form): Cho phép tải ảnh, kịch bản và tệp âm thanh trực tuyến.\n"
        "• Bộ 4 Phiếu học tập số hóa phân hóa: Biên soạn riêng cho 4 nhóm di tích (Bưu điện, Nhà thờ Đức Bà, Chợ Bến Thành, Bến Nhà Rồng) trong tệp văn bản: Phieu_Hoc_Tap_4_Nhom_STEM_Lop8.docx.\n"
        "• Bộ công cụ đánh giá Rubrics: 4 tiêu chí rõ ràng (Bài viết tiếng Anh 30đ, Audio Guide 30đ, Thuyết trình trước lớp 20đ, Tương tác số & Nhóm 20đ)."
    )

    # 6. Khối lớp, Thời lượng, Không gian
    p = doc.add_paragraph()
    r = p.add_run("6. Đối tượng, Thời lượng và Không gian thực hiện:\n"); r.font.bold = True; r.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    p.add_run(
        "• Đối tượng học sinh: Khối 8 (Lớp dạy thực nghiệm: 8A2, sĩ số khoảng 40 - 45 học sinh, chia thành 4 nhóm chuyên trách).\n"
        "• Thời lượng: 02 tiết học trên lớp (90 phút) + trải nghiệm số ngoài giờ lên lớp.\n"
        "• Không gian thực hiện:\n"
        "  - Phòng học thông minh / Phòng học bộ môn Tiếng Anh Lớp 8A2: Bố trí không gian nhóm, máy chiếu/SmartTV và Wi-Fi.\n"
        "  - Không gian gia đình: Góc học tập yên tĩnh để thu âm giọng đọc tiếng Anh đạt chuẩn âm học.\n"
        "  - Không gian số Internet: Cổng học liệu số [english8-two.vercel.app] kết nối toàn cầu, phụ huynh và du khách có thể quét nghe mọi lúc."
    )

    # 7. Khung ký tên & Phê duyệt của Tổ trưởng
    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    tbl_sign = doc.add_table(rows=1, cols=2)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_sign.autofit = False
    tbl_sign.rows[0].cells[0].width = Inches(3.3); tbl_sign.rows[0].cells[1].width = Inches(3.4)

    c0, c1 = tbl_sign.rows[0].cells[0], tbl_sign.rows[0].cells[1]
    p = c0.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Ý KIẾN DUYỆT CỦA TỔ TRƯỞNG CHUYÊN MÔN\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p.add_run("“Đề tài có tính ứng dụng cao, sáng tạo, đúng chuẩn CV 5512.\nNhất trí đề xuất dự thi cấp Thành phố.”\n"); r.font.italic = True; r.font.size = Pt(9.5)
    r = p.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"); r.font.italic = True; r.font.size = Pt(9.5)
    r = p.add_run("[Chữ ký Tổ trưởng chuyên môn]"); r.font.bold = True; r.font.size = Pt(10)

    p = c1.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TP. Hồ Chí Minh, ngày       tháng 11 năm 2026\n"); r.font.italic = True; r.font.size = Pt(10)
    r = p.add_run("GIÁO VIÊN ĐĂNG KÝ CHỦ ĐỀ\n"); r.font.bold = True; r.font.size = Pt(10)
    r = p.add_run("(Ký và ghi rõ họ tên)\n\n\n\n"); r.font.italic = True; r.font.size = Pt(9.5)
    r = p.add_run("Trần Thị Thanh Nguyên"); r.font.bold = True; r.font.size = Pt(10.5)

    # Save to both locations
    dest1 = os.path.join("01_Ho_So_Du_Thi_SGDDT", "Phieu_Dang_Ky_Chu_De_STEM_Lop8.docx")
    dest2 = os.path.join("02_Bao_Cao_Demo_Gui_Khach", "Phieu_Dang_Ky_Chu_De_STEM_Lop8.docx")

    doc.save(dest1)
    doc.save(dest2)
    print(f"Saved: {dest1}")
    print(f"Saved: {dest2}")

if __name__ == "__main__":
    build_registration_docx()
