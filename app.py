import streamlit as st
import google.generativeai as genai
import io
from docx import Document
from pptx import Presentation

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

# Sidebar: Nhập API Key và thông tin bài học
with st.sidebar:
    st.header("⚙️ Thiết lập hệ thống")
    api_key_input = st.text_input("Nhập Google Gemini API Key:", type="password")
    api_key = api_key_input.strip() or st.secrets.get("GEMINI_API_KEY", "").strip()
    
    st.divider()
    st.header("📋 Thông tin bài học")
    mon_hoc = st.selectbox("Môn học:", [
        "Toán học", "Ngữ văn", "Tiếng Anh", "Khoa học tự nhiên", 
        "Vật lí", "Hóa học", "Sinh học", "Lịch sử và Địa lí", 
        "Lịch sử", "Địa lí", "Tin học", "Giáo dục công dân", "Công nghệ"
    ])
    lop = st.selectbox("Khối lớp:", [f"Lớp {i}" for i in range(6, 13)])
    ten_bai = st.text_input("Tên bài học:", value="Hệ phương trình bậc nhất hai ẩn")
    so_tiet = st.number_input("Thời lượng (tiết):", min_value=1, max_value=6, value=2)
    yeu_cau_can_dat = st.text_area(
        "Yêu cầu cần đạt (theo CT GDPT 2018):", 
        height=100, 
        value="- Nhận biết khái niệm hệ phương trình bậc nhất hai ẩn.\n- Giải được hệ phương trình bằng phương pháp thế hoặc cộng đại số."
    )

def call_gemini(prompt_text, key):
    if not key:
        st.error("Vui lòng nhập Google Gemini API Key vào thanh bên trái!")
        return None
    try:
        genai.configure(api_key=key)
        candidate_models = [
            "gemini-1.5-flash-latest",
            "gemini-1.5-flash",
            "gemini-1.5-pro-latest",
            "gemini-1.5-pro",
            "gemini-pro"
        ]
        
        last_error = None
        for model_name in candidate_models:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content(prompt_text)
                if response and response.text:
                    return response.text
            except Exception as err:
                last_error = err
                continue
                
        raise last_error
    except Exception as e:
        st.error(f"Lỗi khi xử lý qua AI: {e}")
        return None

def export_docx(title, content):
    doc = Document()
    doc.add_heading(title, level=0)
    doc.add_paragraph("Biên soạn: Lê Minh Tuấn - GV\nQuy chuẩn: Công văn số 7991/BGDĐT-GDTrH\n")
    for para in content.split("\n"):
        doc.add_paragraph(para)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def export_pptx(title, raw_text):
    prs = Presentation()
    # Slide bìa
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"Môn học: {mon_hoc} - {lop}\nGiáo viên: Lê Minh Tuấn - GV"
    
    # Slides nội dung
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

tab1, tab2, tab3 = st.tabs([
    "📄 1. Kế hoạch bài dạy (CV 7991)", 
    "🖥️ 2. Slide bài giảng (.PPTX)", 
    "📝 3. Đề kiểm tra & Ma trận (CV 7991)"
])

# TAB 1: KẾ HOẠCH BÀI DẠY
with tab1:
    st.subheader("Soạn Kế hoạch bài dạy (Giáo án)")
    if st.button("🚀 Khởi tạo Kế hoạch bài dạy", key="btn_khbd"):
        with st.spinner("Đang biên soạn KHBD theo đúng chuẩn khung văn bản Bộ GD&ĐT..."):
            prompt_khbd = f"""
            Đóng vai trò là chuyên gia sư phạm. Hãy biên soạn Kế hoạch bài dạy chuẩn cho:
            - Môn: {mon_hoc} - {lop}
            - Tên bài: {ten_bai} (Thời lượng: {so_tiet} tiết)
            - Yêu cầu cần đạt: {yeu_cau_can_dat}
            
            Khung kế hoạch bài dạy:
            I. MỤC TIÊU (Năng lực đặc thù, Năng lực chung, Phẩm chất)
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC:
            1. Hoạt động 1: Mở đầu/Khởi động (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện: Chuyển giao -> Thực hiện -> Báo cáo/Thảo luận -> Kết luận/Nhận định)
            2. Hoạt động 2: Hình thành kiến thức mới (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            3. Hoạt động 3: Luyện tập (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            4. Hoạt động 4: Vận dụng (Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện theo 4 bước)
            
            Ký tên người soạn: Lê Minh Tuấn - GV.
            """
            st.session_state["res_khbd"] = call_gemini(prompt_khbd, api_key)
            
    if st.session_state.get("res_khbd"):
        st.markdown(st.session_state["res_khbd"])
        docx_bytes = export_docx(f"KHBD_{ten_bai}", st.session_state["res_khbd"])
        st.download_button(
            "📥 Tải về KHBD Word (.docx)", 
            data=docx_bytes, 
            file_name=f"KHBD_{ten_bai}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# TAB 2: SLIDE BÀI GIẢNG
with tab2:
    st.subheader("Tạo Slide bài giảng")
    so_slide = st.slider("Số lượng Slide:", 5, 15, 8)
    if st.button("🚀 Khởi tạo Slide bài giảng", key="btn_slide"):
        with st.spinner("Đang cấu trúc slide trình chiếu..."):
            prompt_slide = f"""
            Tạo cấu trúc trình chiếu PowerPoint gồm {so_slide} slide cho:
            Bài học: {ten_bai} ({mon_hoc} - {lop}).
            
            ĐỊNH DẠNG BẮT BUỘC:
            Mỗi slide bắt đầu bằng cụm từ: "SLIDE: [Tiêu đề slide]"
            Dưới tiêu đề slide là các gạch đầu dòng súc tích, trình bày phương pháp trực quan (3-4 bullet points/slide).
            Không dùng đoạn văn dài dòng.
            """
            st.session_state["res_slide"] = call_gemini(prompt_slide, api_key)

    if st.session_state.get("res_slide"):
        st.text_area("Nội dung Slide:", value=st.session_state["res_slide"], height=250)
        pptx_bytes = export_pptx(ten_bai, st.session_state["res_slide"])
        st.download_button(
            "📥 Tải về Slide PowerPoint (.pptx)", 
            data=pptx_bytes, 
            file_name=f"Slide_{ten_bai}.pptx", 
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

# TAB 3: ĐỀ KIỂM TRA ĐỊNH KỲ THEO CÔNG VĂN 7991
with tab3:
    st.subheader("Tạo Ma trận, Bản đặc tả và Đề kiểm tra chuẩn CV 7991")
    col1, col2 = st.columns(2)
    with col1:
        thoi_gian = st.selectbox("Thời lượng kiểm tra:", ["45 phút (Định kì)", "60 phút", "90 phút (Cuối kì)"])
    with col2:
        co_tra_loi_ngan = st.checkbox("Môn có trắc nghiệm Trả lời ngắn (Toán, KHTN, Tin...)", value=True)

    if st.button("🚀 Khởi tạo Ma trận & Đề kiểm tra 7991", key="btn_exam_7991"):
        with st.spinner("Đang thiết lập ma trận và đặc tả theo Công văn 7991/BGDĐT-GDTrH..."):
            prompt_7991 = f"""
            Đóng vai trò là chuyên gia khảo thí của Bộ GDĐT. Hãy thiết kế bộ hồ sơ kiểm tra định kì chuẩn xác theo CÔNG VĂN SỐ 7991/BGDĐT-GDTrH:
            - Môn: {mon_hoc} - {lop}
            - Bài / Chủ đề kiểm tra: {ten_bai}
            - Thời gian làm bài: {thoi_gian}
            - Giáo viên thẩm định / ra đề: Lê Minh Tuấn - GV

            QUY CHUẨN CÔNG VĂN 7991:
            1. Tỉ lệ điểm định dạng câu hỏi:
               - Phần I. Trắc nghiệm Nhiều lựa chọn: 3,0 điểm (30%).
               - Phần II. Trắc nghiệm Đúng - Sai: 2,0 điểm (20%). Gồm các câu có 4 ý a, b, c, d (chọn Đúng hoặc Sai).
               {"- Phần III. Trắc nghiệm Trả lời ngắn: 2,0 điểm (20%)." if co_tra_loi_ngan else "- Môn không dùng Trả lời ngắn: Chuyển toàn bộ 2,0 điểm sang dạng Đúng - Sai (Tổng Đúng-Sai 4,0 điểm)."}
               - Phần IV. Tự luận: 3,0 điểm (30%).
            2. Tỉ lệ mức độ nhận thức:
               - Biết: khoảng 40% (4,0 điểm)
               - Hiểu: khoảng 30% (3,0 điểm)
               - Vận dụng: khoảng 30% (3,0 điểm)

            XUẤT RA CHI TIẾT 4 PHẦN:
            1. KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KÌ (Vẽ bảng theo đúng mẫu Phụ lục 1 CV 7991).
            2. BẢN ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KÌ (Vẽ bảng theo đúng mẫu Phụ lục 2 CV 7991: TT, Chủ đề, Nội dung, Yêu cầu cần đạt, Số câu theo mức độ Biết/Hiểu/Vận dụng cho từng định dạng).
            3. ĐỀ KIỂM TRA HOÀN CHỈNH (Trình bày rõ từng phần câu hỏi I, II, III, IV).
            4. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM:
               - Đáp án trắc nghiệm nhiều lựa chọn.
               - Hướng dẫn tính điểm Đúng-Sai: 1 ý đúng = 0.1đ; 2 ý đúng = 0.25đ; 3 ý đúng = 0.5đ; 4 ý đúng = 1.0đ.
               - Đáp án phần trả lời ngắn (nếu có).
               - Barem điểm chi tiết cho bài Tự luận.
            """
            st.session_state["res_exam"] = call_gemini(prompt_7991, api_key)
            
    if st.session_state.get("res_exam"):
        st.markdown(st.session_state["res_exam"])
        docx_exam_bytes = export_docx(f"De_Kiem_Tra_7991_{ten_bai}", st.session_state["res_exam"])
        st.download_button(
            "📥 Tải về Bộ Đề kiểm tra chuẩn CV 7991 (.docx)", 
            data=docx_exam_bytes, 
            file_name=f"De_Kiem_Tra_7991_{ten_bai}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
