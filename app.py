import streamlit as st
import google.generativeai as genai
import io
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# Cấu hình trang
st.set_page_config(
    page_title="AI Giáo Dục 7991 - Lê Minh Tuấn",
    page_icon="🎓",
    layout="wide"
)

# Header định danh tác giả
st.markdown("## 🎓 ỨNG DỤNG AI HỖ TRỢ SOẠN GIẢNG THEO CÔNG VĂN 7991")
st.markdown("**Tác giả:** Lê Minh Tuấn - GV")
st.caption("Chương trình chuyển giao kỹ thuật ứng dụng AI trong GD – Tỉnh Vĩnh Long (10/2026)")
st.divider()

# Sidebar: Cấu hình API và Thông tin bài dạy
with st.sidebar:
    st.header("⚙️ Thiết lập hệ thống")
    # Lấy key từ secrets nếu có, hoặc cho nhập thủ công
    api_key_input = st.text_input("Nhập Google Gemini API Key:", type="password")
    api_key = api_key_input or st.secrets.get("GEMINI_API_KEY", "")
    
    st.divider()
    st.header("📋 Thông tin bài học")
    mon_hoc = st.selectbox("Môn học:", ["Toán học", "Ngữ văn", "Tiếng Anh", "KHTN / Vật lý / Hóa học / Sinh học", "Lịch sử - Địa lý", "Tin học", "GDCD", "Công nghệ"])
    lop = st.selectbox("Khối lớp:", [f"Lớp {i}" for i in range(6, 13)])
    ten_bai = st.text_input("Tên bài học:", value="Hệ phương trình bậc nhất hai ẩn")
    so_tiet = st.number_input("Thời lượng (tiết):", min_value=1, max_value=6, value=2)
    yeu_cau_can_dat = st.text_area("Yêu cầu cần đạt (theo CT GDPT 2018):", height=90, 
                                  value="- Nhận biết khái niệm hệ phương trình bậc nhất hai ẩn.\n- Giải được hệ phương trình bằng phương pháp thế hoặc cộng đại số.")

# Hàm gọi AI Gemini
def call_gemini(prompt_text, api_key):
    if not api_key:
        st.error("Vui lòng nhập API Key ở thanh bên trái!")
        return None
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")
        response = model.generate_content(prompt_text)
        return response.text
    except Exception as e:
        st.error(f"Lỗi khi xử lý qua AI: {e}")
        return None

# Hàm tạo file Word
def export_docx(title, content):
    doc = Document()
    doc.add_heading(title, level=0)
    doc.add_paragraph("Biên soạn bởi: Lê Minh Tuấn - GV\nCấu trúc theo chuẩn Công văn 7991/BGDĐT\n")
    for para in content.split("\n"):
        doc.add_paragraph(para)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Hàm tạo file PowerPoint
def export_pptx(title, raw_text):
    prs = Presentation()
    # Slide bìa
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"Môn học: {mon_hoc} - {lop}\nNgười soạn: Lê Minh Tuấn - GV"
    
    # Các slide nội dung phân cách bằng "SLIDE:"
    slides_data = [s for s in raw_text.split("SLIDE:") if s.strip()]
    for s_text in slides_data:
        lines = [line.strip() for line in s_text.strip().split("\n") if line.strip()]
        if not lines:
            continue
        slide_title = lines[0].replace("#", "").strip()
        body_lines = lines[1:]
        
        slide_layout = prs.slide_layouts[1]
        new_slide = prs.slides.add_slide(slide_layout)
        new_slide.shapes.title.text = slide_title
        
        body_shape = new_slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        for b_line in body_lines:
            p = tf.add_paragraph()
            p.text = b_line.lstrip("-*• ")
            p.level = 0
            
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

# Giao diện Tabs tác vụ
tab1, tab2, tab3 = st.tabs([
    "📄 1. Kế hoạch bài dạy (CV 7991)", 
    "🖥️ 2. Slide bài giảng (.PPTX)", 
    "📝 3. Đề kiểm tra & Ma trận"
])

# TAB 1: KẾ HOẠCH BÀI DẠY
with tab1:
    st.subheader("Soạn Kế hoạch bài dạy (Giáo án)")
    if st.button("🚀 Khởi tạo Kế hoạch bài dạy", key="btn_khbd"):
        with st.spinner("AI đang thiết kế theo khung Công văn 7991..."):
            prompt = f"""
            Đóng vai trò là chuyên gia sư phạm. Hãy biên soạn Kế hoạch bài dạy chi tiết cho:
            - Môn: {mon_hoc} - {lop}
            - Tên bài: {ten_bai} (Thời lượng: {so_tiet} tiết)
            - Yêu cầu cần đạt: {yeu_cau_can_dat}
            
            Khung kế hoạch bài dạy bắt buộc tuân theo Công văn 7991/BGDĐT:
            I. MỤC TIÊU (Năng lực đặc thù, Năng lực chung, Phẩm chất)
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (Của giáo viên và học sinh)
            III. TIẾN TRÌNH DẠY HỌC:
            Chia rõ 4 hoạt động:
            1. Hoạt động 1: Mở đầu/Khởi động (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện: Giao nhiệm vụ -> Thực hiện -> Báo cáo/Thảo luận -> Kết luận/Nhận định)
            2. Hoạt động 2: Hình thành kiến thức mới (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            3. Hoạt động 3: Luyện tập (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            4. Hoạt động 4: Vận dụng (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            Người soạn: Lê Minh Tuấn - GV.
            """
            st.session_state["res_khbd"] = call_gemini(prompt, api_key)
            
    if "res_khbd" in st.session_state and st.session_state["res_khbd"]:
        st.markdown(st.session_state["res_khbd"])
        docx_bytes = export_docx(f"KHBD_{ten_bai}", st.session_state["res_khbd"])
        st.download_button("📥 Tải về file Word (.docx)", data=docx_bytes, file_name=f"KHBD_{ten_bai}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# TAB 2: SLIDE BÀI GIẢNG
with tab2:
    st.subheader("Tạo dàn ý và xuất Slide bài giảng")
    so_slide = st.slider("Số lượng Slide dự kiến:", 5, 15, 8)
    if st.button("🚀 Khởi tạo Slide bài giảng", key="btn_slide"):
        with st.spinner("Đang xây dựng nội dung bài giảng đa phương tiện..."):
            prompt = f"""
            Tạo cấu trúc trình chiếu bài giảng PowerPoint gồm {so_slide} slide cho:
            Bài học: {ten_bai} ({mon_hoc} - {lop}).
            
            Yêu cầu định dạng CHÍNH XÁC:
            Mỗi slide bắt đầu bằng chữ "SLIDE: [Tiêu đề slide]"
            Dưới tiêu đề slide là các dấu gạch đầu dòng ngắn gọn súc tích (3-4 ý/slide).
            Không dùng văn xuôi dài dòng.
            """
            st.session_state["res_slide"] = call_gemini(prompt, api_key)

    if "res_slide" in st.session_state and st.session_state["res_slide"]:
        st.text_area("Nội dung Slide:", value=st.session_state["res_slide"], height=250)
        pptx_bytes = export_pptx(ten_bai, st.session_state["res_slide"])
        st.download_button("📥 Tải về file PowerPoint (.pptx)", data=pptx_bytes, file_name=f"Slide_{ten_bai}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# TAB 3: ĐỀ KIỂM TRA THEO CÔNG VĂN 7991/BGDĐT-GDTrH
with tab3:
    st.subheader("Tạo Ma trận, Bản đặc tả và Đề kiểm tra chuẩn CV 7991")
    col1, col2 = st.columns(2)
    with col1:
        mon_thi = st.text_input("Môn kiểm tra:", value=mon_hoc)
    with col2:
        thoi_gian = st.selectbox("Thời lượng kiểm tra:", ["45 phút", "60 phút", "90 phút"])
        
    co_tra_loi_ngan = st.checkbox("Môn có sử dụng câu hỏi Trả lời ngắn (Toán, KHTN, Tin...)", value=True)

    if st.button("🚀 Tạo Đề & Ma trận theo CV 7991", key="btn_exam_7991"):
        with st.spinner("Đang xây dựng theo cấu trúc ma trận chuẩn Công văn 7991..."):
            prompt_7991 = f"""
            Đóng vai trò là chuyên gia khảo thí Bộ GDĐT. Hãy thiết kế bộ tài liệu kiểm tra định kỳ môn {mon_thi} - {lop} cho bài/chủ đề: {ten_bai}.
            Thời gian làm bài: {thoi_gian}.
            Người thẩm định và biên soạn: Lê Minh Tuấn - GV.

            BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT CÔNG VĂN SỐ 7991/BGDĐT-GDTrH:
            1. TỈ LỆ ĐIỂM VÀ ĐỊNH DẠNG:
               - Phần 1. Trắc nghiệm nhiều lựa chọn: khoảng 3,0 điểm (30%).
               - Phần 2. Trắc nghiệm Đúng - Sai: khoảng 2,0 điểm (20%). Mỗi câu gồm 4 lệnh hỏi a, b, c, d (chọn Đúng hoặc Sai).
               {"- Phần 3. Trắc nghiệm Trả lời ngắn: khoảng 2,0 điểm (20%)." if co_tra_loi_ngan else "- Chuyển 2.0 điểm của Trả lời ngắn sang dạng Đúng - Sai theo chú thích 3 của CV 7991."}
               - Phần 4. Tự luận: khoảng 3,0 điểm (30%).
               - Tỉ lệ mức độ đánh giá: Biết: 40% (4,0 điểm), Hiểu: 30% (3,0 điểm), Vận dụng: 30% (3,0 điểm).

            2. NỘI DUNG XUẤT RA GỒM ĐỦ 4 PHẦN:
               I. KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KÌ (Vẽ bảng Markdown theo mẫu Phụ lục 1 CV 7991).
               II. BẢN ĐẶC TẢ ĐỀ KIỂM TRA (Vẽ bảng theo Phụ lục 2 CV 7991, có cột Yêu cầu cần đạt và Năng lực).
               III. ĐỀ KIỂM TRA ĐỊNH KỲ HOÀN CHỈNH (Trình bày rõ ràng 4 phần câu hỏi).
               IV. HƯỚNG DẪN CHẤM VÀ ĐÁP ÁN (Gồm quy tắc tính điểm Đúng - Sai theo chuẩn: 1 ý=0.1đ, 2 ý=0.25đ, 3 ý=0.5đ, 4 ý=1.0đ; đáp án trả lời ngắn và barem điểm tự luận chi tiết).
            """
            st.session_state["res_exam"] = call_gemini(prompt_7991, api_key)
            
    if "res_exam" in st.session_state and st.session_state["res_exam"]:
        st.markdown(st.session_state["res_exam"])
        docx_exam_bytes = export_docx(f"De_Kiem_Tra_7991_{ten_bai}", st.session_state["res_exam"])
        st.download_button(
            "📥 Tải về Hồ sơ Đề kiểm tra chuẩn CV 7991 (.docx)", 
            data=docx_exam_bytes, 
            file_name=f"De_Kiem_Tra_7991_{ten_bai}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )