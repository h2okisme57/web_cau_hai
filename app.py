import io
import re
import time
import zipfile
import xml.etree.ElementTree as ET
import streamlit as st
from google import genai
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

# Danh sách các model có quota độc lập, ưu tiên các model chưa bị hết lượt
MODEL_CANDIDATES = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-2.5-flash",
    "gemini-2.0-flash-exp",
]

st.set_page_config(
    page_title="AI Giáo Dục - Lê Minh Tuấn",
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
    
    # Nút kiểm tra API Key tự động dò model còn quota
    if st.button("🔍 Kiểm tra API Key", use_container_width=True):
        if not api_key:
            st.error("Chưa nhập API Key!")
        else:
            with st.spinner("Đang tìm model còn quota khả dụng..."):
                client = genai.Client(api_key=api_key)
                connected_model = None
                last_err = None
                
                # Thử lần lượt các model để tìm model chưa chạm trần 429
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
                    st.success(f"✅ Kết nối thành công!\n\nĐang sử dụng: `{connected_model}`")
                else:
                    st.error(f"❌ Tất cả model trong danh sách đều hết quota hoặc lỗi: {last_err}")

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

# Hàm đọc file docx an toàn có fallback giải nén XML
def read_uploaded_file(uploaded_file):
    try:
        raw_bytes = uploaded_file.read()
        if uploaded_file.name.endswith(".docx"):
            # Cách 1: Thử bằng thư viện docx chuẩn
            try:
                doc = Document(io.BytesIO(raw_bytes))
                full_text = []
                for p in doc.paragraphs:
                    if p.text.strip():
                        full_text.append(p.text.strip())
                for table in doc.tables:
                    for row in table.rows:
                        row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if row_text:
                            full_text.append(" | ".join(row_text))
                extracted = "\n".join(full_text)
                if extracted.strip():
                    return extracted
            except Exception:
                pass

            # Cách 2: Mở trực tiếp file ZIP XML nếu header docx bị sai lệch
            try:
                with zipfile.ZipFile(io.BytesIO(raw_bytes)) as z:
                    target_xml = None
                    for name in z.namelist():
                        if name.endswith("word/document.xml") or name.endswith("document.xml"):
                            target_xml = name
                            break
                    if target_xml:
                        xml_content = z.read(target_xml)
                        tree = ET.fromstring(xml_content)
                        texts = [elem.text for elem in tree.iter() if elem.tag.endswith('t') and elem.text]
                        return "\n".join(texts)
            except Exception:
                pass

        elif uploaded_file.name.endswith(".txt"):
            return raw_bytes.decode("utf-8", errors="ignore")

    except Exception as e:
        st.error(f"Lỗi khi đọc file tài liệu: {e}")
        return ""
    return ""

# Hàm gọi Gemini AI tự động né lỗi 429 và 503
def call_gemini(prompt_text, key):
    if not key:
        st.error("Vui lòng nhập API Key ở thanh bên trái!")
        return None
    try:
        client = genai.Client(api_key=key.strip())
        last_error = None
        
        # Sắp xếp thứ tự thử nghiệm: ưu tiên model đang active trước
        active_m = st.session_state.get("active_model")
        trial_models = [active_m] + [m for m in MODEL_CANDIDATES if m != active_m] if active_m else MODEL_CANDIDATES

        for model_name in trial_models:
            for retry in range(2):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt_text,
                    )
                    
                    if response and hasattr(response, "text") and response.text:
                        st.session_state["active_model"] = model_name
                        return response.text
                    
                    if response and response.candidates:
                        parts_text = "".join([p.text for p in response.candidates[0].content.parts if hasattr(p, "text") and p.text])
