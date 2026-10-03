# 1. ĐOẠN XỬ LÝ TRONG NÚT "Kiểm tra API Key":
    if st.button("🔍 Kiểm tra API Key", use_container_width=True):
        if not api_key:
            st.error("Chưa nhập API Key!")
        else:
            with st.spinner("Đang xác thực và chọn model tối ưu..."):
                try:
                    client = genai.Client(api_key=api_key)
                    supported_models = [m.name for m in client.models.list()]
                    
                    # Ưu tiên các model thế hệ mới không bị deprecate
                    preferred_order = [
                        "gemini-3.8-flash",
                        "gemini-3-flash",
                        "gemini-2.0-flash",
                        "gemini-1.5-flash-latest",
                        "gemini-1.5-flash",
                        "gemini-1.5-pro"
                    ]
                    
                    best_model = None
                    for target in preferred_order:
                        for m_name in supported_models:
                            if target in m_name:
                                best_model = m_name
                                break
                        if best_model:
                            break
                            
                    target_model = best_model if best_model else (supported_models[0] if supported_models else "gemini-3.8-flash")
                    st.session_state["active_model"] = target_model
                    st.success(f"✅ Kết nối thành công!\n\nModel sẵn sàng: `{target_model}`")
                except Exception as e:
                    st.error(f"❌ Lỗi xác thực Key: {e}")

# 2. ĐOẠN HÀM call_gemini:
def call_gemini(prompt_text, key):
    if not key:
        st.error("Vui lòng nhập API Key ở thanh bên trái!")
        return None
    try:
        client = genai.Client(api_key=key)
        
        target_model = st.session_state.get("active_model")
        if not target_model:
            preferred_order = [
                "gemini-3.8-flash",
                "gemini-3-flash",
                "gemini-2.0-flash",
                "gemini-1.5-flash-latest",
                "gemini-1.5-flash",
                "gemini-1.5-pro"
            ]
            try:
                available = [m.name for m in client.models.list()]
                for target in preferred_order:
                    for m_name in available:
                        if target in m_name:
                            target_model = m_name
                            break
                    if target_model:
                        break
            except Exception:
                target_model = "gemini-3.8-flash"
                
            if not target_model:
                target_model = "gemini-3.8-flash"
            st.session_state["active_model"] = target_model

        response = client.models.generate_content(
            model=target_model,
            contents=prompt_text,
        )
        
        if response and response.text:
            return response.text
        else:
            st.warning("Không có phản hồi từ mô hình.")
            return None
    except Exception as e:
        st.error(f"Lỗi khi xử lý qua AI: {e}")
        return None
