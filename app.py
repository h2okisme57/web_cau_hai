import io
import time
import streamlit as st
from google import genai
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# Danh sách model dự phòng theo thứ tự ưu tiên nhằm tránh lỗi 503 / 404
MODEL_CANDIDATES = [
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-3.8-flash",
    "gemini-1.5-pro",
]

st.set_page_config(
    page_title="AI Giáo Dục 7991 - Lê Minh Tuấn",
    page_icon="🎓",
    layout="wide"
)

# Header định danh tác giả
st.markdown("## 🎓 HỆ THỐNG TRỢ LÝ AI GIÁO DỤC TOÀN DIỆN")
st.markdown("**Tác giả:** Lê Minh Tuấn - GV (Chuyên môn: Lịch sử & Địa lí)")
st.caption("Chương trình chuyển giao kỹ thuật ứng dụng AI trong GD – Tỉnh Vĩnh Long (10/2026)")
st.divider()

# Sidebar: Thiết lập API & Thông tin bài học
with st.sidebar:
    st.header("⚙️ Thiết lập hệ thống")
    api_key_input = st.text_input("Nhập Google Gemini API Key:", type="password")
    api_key = api_key_input.strip() or st.secrets.get("GEMINI_API_KEY", "").strip()
    
    # Nút kiểm tra API Key với cơ chế quét fallback né lỗi 503
    if st.button("🔍 Kiểm tra API Key", use_container_width=True):
        if not api_key:
            st.error("Chưa nhập API Key!")
        else:
            with st.spinner("Đang ping kiểm tra kết nối với hệ thống Google AI..."):
                client = genai.Client(api_key=api_key)
                connected_model = None
                last_err = None
                
                for candidate in MODEL_CANDIDATES:
                    try:
                        res = client.models.generate_content(
                            model=candidate,
                            contents="ping"
                        )
                        if res:
                            connected_model = candidate
                            break
                    except Exception as err:
                        last_err = err
                        continue
                
                if connected_model:
                    st.session_state["active_model"] = connected_model
                    st.success(f"✅ Kết nối thành công!\n\nModel sẵn sàng: `{connected_model}`")
                else:
                    st.error(f"❌ Kết nối thất bại: {last_err}")

    st.divider()
    st.header("📋 Thông tin bài dạy thực nghiệm")
    mon_hoc = st.selectbox("Môn học:", [
        "Lịch sử và Địa lí (Lịch sử)", "Lịch sử", "Địa lí", "Toán học", 
        "Ngữ văn", "Khoa học tự nhiên", "Tin học", "Giáo dục công dân"
    ], index=0)
    lop = st.selectbox("Khối lớp:", [f"Lớp {i}" for i in range(6, 13)], index=1)
    ten_bai = st.text_input("Tên bài học:", value="Bài 2. Các cuộc phát kiến địa lí")
    so_tiet = st.number_input("Thời lượng (tiết):", min_value=1, max_value=6, value=2)
    yeu_cau_can_dat = st.text_area(
        "Yêu cầu cần đạt (theo CT GDPT 2018):", 
        height=100, 
        value="- Trình bày được nguyên nhân và điều kiện của các cuộc phát kiến địa lí.\n- Mô tả được các cuộc phát kiến địa lí của B. Đi-a-xơ, C. Cô-lôm-bô, V. Ga-ma và Ph. Ma-gien-lan.\n- Đánh giá được tác động của các cuộc phát kiến địa lí đối với tiến trình lịch sử."
    )

# Hàm gọi Gemini AI với cơ chế tự động thử lại và đổi model nếu quá tải (503)
def call_gemini(prompt_text, key):
    if not key:
        st.error("Vui lòng nhập API Key ở thanh bên trái!")
        return None
    try:
        client = genai.Client(api_key=key)
        last_error = None
        
        # Nếu đã có model test thành công trước đó thì ưu tiên dùng trước
        active_model = st.session_state.get("active_model")
        trial_list = [active_model] + [m for m in MODEL_CANDIDATES if m != active_model] if active_model else MODEL_CANDIDATES

        for model_name in trial_list:
            for retry in range(2):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt_text,
                    )
                    if response and response.text:
                        st.session_state["active_model"] = model_name
                        return response.text
                except Exception as err:
                    last_error = err
                    err_msg = str(err).lower()
                    if "503" in err_msg or "high demand" in err_msg or "unavailable" in err_msg:
                        time.sleep(1)
                        continue
                    else:
                        break
                        
        st.error(f"Lỗi khi xử lý qua AI: {last_error}")
        return None
    except Exception as e:
        st.error(f"Lỗi kết nối client AI: {e}")
        return None

# Xuất file Word (.docx)
def export_docx(title, content, author="Lê Minh Tuấn - GV"):
    doc = Document()
    doc.add_heading(title, level=0)
    p = doc.add_paragraph()
    p.add_run(f"Người thực hiện: {author}\n").bold = True
    p.add_run("Quy chuẩn chuyên môn: Công văn 5512/BGDĐT & Công văn 7991/BGDĐT-GDTrH\n").italic = True
    for para in content.split("\n"):
        doc.add_paragraph(para)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Hàm xuất PowerPoint chuẩn Bài tập 3 (TO - VỪA - NHỎ + Speaker Notes)
def export_advanced_pptx(title, raw_text, author):
    prs = Presentation()
    prs.slide_width = Inches(13.333) # 16:9 widescreen
    prs.slide_height = Inches(7.5)
    
    # Slide 1: Bìa
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"Môn học: {mon_hoc} - {lop}\nGiáo viên biên soạn: {author}"
    
    # Tách và định dạng các slide nội dung
    slides_raw = [s for s in raw_text.split("SLIDE:") if s.strip()]
    for s_item in slides_raw:
        lines = [l.strip() for l in s_item.strip().split("\n") if l.strip()]
        if not lines:
            continue
        slide_title = lines[0].replace("#", "").strip()
        body_lines = []
        speaker_notes = []
        is_notes = False
        
        for l in lines[1:]:
            if "LỜI DẪN GV:" in l.upper() or "GHI CHÚ GV:" in l.upper() or "NOTES:" in l.upper():
                is_notes = True
                continue
            if is_notes:
                speaker_notes.append(l)
            else:
                body_lines.append(l)
                
        # Tạo slide nội dung
        new_slide = prs.slides.add_slide(prs.slide_layouts[1])
        new_slide.shapes.title.text = slide_title
        
        # Đổ nội dung vào text frame
        tf = new_slide.placeholders[1].text_frame
        tf.clear()
        for bl in body_lines:
            p = tf.add_paragraph()
            clean_text = bl.lstrip("-*• ")
            if clean_text.startswith("TO:"):
                p.text = clean_text.replace("TO:", "").strip()
                p.font.size = Pt(26)
                p.font.bold = True
            elif clean_text.startswith("VỪA:"):
                p.text = clean_text.replace("VỪA:", "").strip()
                p.font.size = Pt(20)
            elif clean_text.startswith("NHỎ:"):
                p.text = clean_text.replace("NHỎ:", "").strip()
                p.font.size = Pt(16)
                p.font.italic = True
            else:
                p.text = clean_text
                p.font.size = Pt(20)
                
        # Tích hợp lời dẫn giáo viên vào Speaker Notes
        if speaker_notes:
            notes_slide = new_slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = "\n".join(speaker_notes)
            
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

# Giao diện 4 Tabs nghiệp vụ
tab1, tab2, tab3, tab4 = st.tabs([
    "📄 1. Kế hoạch bài dạy (Bài tập 1)", 
    "📝 2. Đề kiểm tra CV 7991 (Bài tập 1)",
    "💡 3. Trợ lý Sáng kiến kinh nghiệm (Bài tập 2)",
    "🖥️ 4. Chuyển KHBD thành Slide PPTX (Bài tập 3)"
])

# TAB 1: KẾ HOẠCH BÀI DẠY (CV 5512)
with tab1:
    st.subheader("Soạn Kế hoạch bài dạy (Giáo án CV 5512)")
    if st.button("🚀 Khởi tạo Kế hoạch bài dạy", key="btn_khbd"):
        with st.spinner("Đang biên soạn KHBD theo đúng chuẩn khung văn bản Bộ GD&ĐT..."):
            prompt_khbd = f"""
            Đóng vai trò là chuyên gia sư phạm. Hãy biên soạn Kế hoạch bài dạy chuẩn cho:
            - Môn: {mon_hoc} - {lop}
            - Tên bài: {ten_bai} (Thời lượng: {so_tiet} tiết)
            - Yêu cầu cần đạt: {yeu_cau_can_dat}
            
            Khung kế hoạch bài dạy chuẩn:
            I. MỤC TIÊU (Năng lực đặc thù, Năng lực chung, Phẩm chất)
            II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
            III. TIẾN TRÌNH DẠY HỌC (Gồm 4 hoạt động: Khởi động, Hình thành kiến thức, Luyện tập, Vận dụng; mỗi hoạt động tổ chức đủ 4 bước: Giao nhiệm vụ -> Thực hiện -> Báo cáo -> Nhận định).
            Ký tên người soạn: Lê Minh Tuấn - GV.
            """
            st.session_state["res_khbd"] = call_gemini(prompt_khbd, api_key)
            
    if st.session_state.get("res_khbd"):
        st.markdown(st.session_state["res_khbd"])
        docx_bytes = export_docx(f"KHBD_{ten_bai}", st.session_state["res_khbd"], "Lê Minh Tuấn - GV")
        st.download_button(
            "📥 Tải về KHBD Word (.docx)", 
            data=docx_bytes, 
            file_name=f"KHBD_LeMinhTuan.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# TAB 2: ĐỀ KIỂM TRA ĐỊNH KỲ THEO CÔNG VĂN 7991
with tab2:
    st.subheader("Tạo Ma trận, Bản đặc tả và Đề kiểm tra chuẩn CV 7991")
    col1, col2 = st.columns(2)
    with col1:
        thoi_gian = st.selectbox("Thời lượng kiểm tra:", ["45 phút (Định kì)", "60 phút", "90 phút"])
    with col2:
        co_tra_loi_ngan = st.checkbox("Môn có trắc nghiệm Trả lời ngắn", value=False)

    if st.button("🚀 Khởi tạo Ma trận & Đề kiểm tra 7991", key="btn_exam_7991"):
        with st.spinner("Đang thiết lập ma trận và đặc tả theo Công văn 7991/BGDĐT-GDTrH..."):
            prompt_7991 = f"""
            Thiết kế bộ hồ sơ kiểm tra định kì chuẩn xác theo CÔNG VĂN SỐ 7991/BGDĐT-GDTrH:
            - Môn: {mon_hoc} - {lop} | Bài: {ten_bai} | Thời gian: {thoi_gian}
            - Giáo viên ra đề: Lê Minh Tuấn - GV
            - Tỉ lệ điểm: Nhiều lựa chọn 30%, Đúng - Sai 20%, {"Trả lời ngắn 20%" if co_tra_loi_ngan else "Chuyển trả lời ngắn sang Đúng-Sai 40%"}, Tự luận 30%.
            - Tỉ lệ nhận thức: Biết 40%, Hiểu 30%, Vận dụng 30%.
            Xuất đầy đủ 4 phần:
            I. KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KÌ (Bảng Markdown).
            II. BẢN ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KÌ (Có Yêu cầu cần đạt và Năng lực).
            III. ĐỀ KIỂM TRA HOÀN CHỈNH.
            IV. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM (Barem điểm Đúng - Sai chuẩn: đúng 1 ý=0.1đ, 2 ý=0.25đ, 3 ý=0.5đ, 4 ý=1.0đ; barem tự luận chi tiết).
            """
            st.session_state["res_exam"] = call_gemini(prompt_7991, api_key)
            
    if st.session_state.get("res_exam"):
        st.markdown(st.session_state["res_exam"])
        docx_exam_bytes = export_docx(f"De_Kiem_Tra_7991_{ten_bai}", st.session_state["res_exam"], "Lê Minh Tuấn - GV")
        st.download_button(
            "📥 Tải về Bộ Đề kiểm tra chuẩn CV 7991 (.docx)", 
            data=docx_exam_bytes, 
            file_name=f"De_Kiem_Tra_7991_LeMinhTuan.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# TAB 3: TRỢ LÝ SÁNG KIẾN KINH NGHIỆM (BÀI TẬP 2)
with tab3:
    st.subheader("💡 Trợ lý Xây dựng Sáng kiến kinh nghiệm (SKKN)")
    ten_skkn = st.text_input("Tên đề tài SKKN:", value="Ứng dụng Trí tuệ nhân tạo (AI) và Bản đồ số hóa trong đổi mới dạy học phân môn Lịch sử cấp THCS")
    thuc_trang = st.text_area("Thực trạng trước khi áp dụng:", value="Học sinh học lịch sử thụ động, việc ghi nhớ các mốc niên đại và hải trình thám hiểm còn máy móc, trừu tượng.")
    if st.button("🚀 Khởi tạo Dự thảo SKKN hoàn chỉnh", key="btn_skkn"):
        with st.spinner("Đang xây dựng dự thảo SKKN theo thể thức chuẩn Thông tư 18/2013/TT-BKHCN..."):
            prompt_skkn = f"""
            Chấp bút một bản DỰ THẢO SÁNG KIẾN KINH NGHIỆM hoàn chỉnh theo quy định mẫu của Bộ GDĐT:
            - Tên sáng kiến: {ten_skkn}
            - Tác giả: Lê Minh Tuấn - GV | Môn: {mon_hoc}
            - Thực trạng: {thuc_trang}
            Gồm đủ: Đơn yêu cầu công nhận sáng kiến (Phụ lục I Thông tư 18), Lý do chọn đề tài, Mục đích, Giải pháp thực hiện chi tiết, Hiệu quả thu được (bảng đối chứng số liệu) và Bài học kinh nghiệm.
            """
            st.session_state["res_skkn"] = call_gemini(prompt_skkn, api_key)
            
    if st.session_state.get("res_skkn"):
        st.markdown(st.session_state["res_skkn"])
        docx_skkn = export_docx(f"SKKN_{ten_skkn[:30]}", st.session_state["res_skkn"], "Lê Minh Tuấn - GV")
        st.download_button(
            "📥 Tải về Bản Dự thảo SKKN (.docx)", 
            data=docx_skkn, 
            file_name="Du_Thao_SKKN_LeMinhTuan.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# TAB 4: CHUYỂN KHBD THÀNH SLIDE POWERPOINT (BÀI TẬP 3 CHUYÊN SÂU)
with tab4:
    st.subheader("🖥️ Trợ lý Chuyển đổi KHBD thành Bộ Slide PowerPoint (.PPTX)")
    st.caption("Thiết kế theo phân cấp TO - VỪA - NHỎ, tích hợp Lời dẫn GV và xuất file .pptx thực tế")
    
    source_choice = st.radio("Chọn nguồn Kế hoạch bài dạy:", ["Sử dụng KHBD đã tạo ở Tab 1", "Nhập / Dán nội dung KHBD thủ công"], horizontal=True)
    
    khbd_input_text = ""
    if source_choice == "Sử dụng KHBD đã tạo ở Tab 1":
        khbd_input_text = st.session_state.get("res_khbd", "")
        if not khbd_input_text:
            st.info("💡 Bạn chưa bấm tạo ở Tab 1, hệ thống sẽ sử dụng thông tin tóm tắt bài học bên trái.")
            khbd_input_text = f"Môn: {mon_hoc} - {lop}. Bài học: {ten_bai}. Thời lượng: {so_tiet} tiết. Yêu cầu cần đạt: {yeu_cau_can_dat}"
    else:
        khbd_input_text = st.text_area("Dán nội dung KHBD vào đây:", height=200, value=f"Môn: {mon_hoc} - {lop}. Bài học: {ten_bai}. Thời lượng: {so_tiet} tiết. Yêu cầu cần đạt: {yeu_cau_can_dat}")

    so_slide_target = st.slider("Số lượng Slide mục tiêu:", 6, 16, 8)
    
    if st.button("🚀 Chuyển đổi KHBD thành Bộ Slide PPTX", type="primary", key="btn_run_bt3"):
        with st.spinner("AI đang phân tích tiến trình KHBD và thiết kế kịch bản trình chiếu phân cấp..."):
            prompt_bt3 = f"""
            ĐÓNG VAI TRÒ: Trợ lý thiết kế nội dung slide dạy học THCS–THPT Việt Nam chuyên nghiệp.
            NHIỆM VỤ: Chuyển toàn bộ Kế hoạch bài dạy dưới đây thành kịch bản trình chiếu PowerPoint gồm {so_slide_target} slide.
            NGƯỜI SOẠN: Lê Minh Tuấn - GV

            NỘI DUNG KẾ HOẠCH BÀI DẠY:
            \"\"\"{khbd_input_text}\"\"\"

            QUY CHUẨN KỊCH BẢN BẮT BUỘC (TUÂN THỦ 100%):
            1. Mỗi slide bắt đầu bằng: "SLIDE: [Tiêu đề ngắn gọn]"
            2. Nội dung hiển thị trên mặt slide theo cấu trúc phân cấp:
               - TO: [Câu hỏi, thông điệp hoặc từ khóa trọng tâm - viết hoa/chữ đậm]
               - VỪA: [3-4 ý chính, nhiệm vụ hoặc dữ kiện thực hiện]
               - NHỎ: [Gợi ý hình ảnh tư liệu, nguồn hoặc hướng dẫn phụ]
            3. Dưới mỗi slide BẮT BUỘC có mục: "LỜI DẪN GV: [2-3 câu giáo viên nói để kết nối và giảng giải, không hiển thị trên slide]"
            4. Tách biệt hoàn toàn phần chữ cho học sinh và lời dẫn cho giáo viên. Không đưa đáp án bài tập trực tiếp lên slide giao việc.
            """
            st.session_state["res_slide_bt3"] = call_gemini(prompt_bt3, api_key)

    if st.session_state.get("res_slide_bt3"):
        st.markdown(st.session_state["res_slide_bt3"])
        pptx_bt3_bytes = export_advanced_pptx(ten_bai, st.session_state["res_slide_bt3"], "Lê Minh Tuấn - GV")
        st.download_button(
            "📥 Tải về Tệp PowerPoint Slide Bài giảng (.pptx)", 
            data=pptx_bt3_bytes, 
            file_name=f"Slide_{ten_bai}_LeMinhTuan.pptx", 
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
