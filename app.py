import streamlit as st
import docx
import openpyxl
import random
import re

# Cấu hình giao diện trang web
st.set_page_config(page_title="Hệ Thống Trắc Nghiệm Tự Động", layout="wide")

st.title("📝 Hệ Thống Tạo & Làm Đề Thi Trắc Nghiệm")
st.write("Tải lên file Word (.docx) hoặc Excel (.xlsx) để tự động trích xuất ngân hàng câu hỏi.")

# Khởi tạo bộ nhớ lưu trữ phiên làm việc (Session State)
if "questions_bank" not in st.session_state:
    st.session_state["questions_bank"] = []
if "current_exam" not in st.session_state:
    st.session_state["current_exam"] = []

# --- HÀM 1: XỬ LÝ FILE WORD (.DOCX) ---
def parse_docx(file):
    doc = docx.Document(file)
    questions = []
    current_q = None
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        # Nhận diện đầu câu hỏi (Ví dụ: Question 1, Câu 1,...)
        if re.match(r'^(Question|\bCâ​u\b|\bCâu\b)\s*\d+', text, re.IGNORECASE):
            if current_q and len(current_q['options']) > 0:
                questions.append(current_q)
            current_q = {"question": text, "options": [], "correct": None}
            
        elif current_q is not None and re.match(r'^[A-D][\.\:\)]', text):
            # Kiểm tra xem dòng/đoạn này có chữ màu xanh hoặc highlight không
            is_correct = False
            for run in p.runs:
                # Kiểm tra màu chữ (RGB)
                if run.font.color and run.font.color.rgb:
                    color_hex = str(run.font.color.rgb)
                    # Mã màu xanh lam/dương (thường bắt đầu bằng 00, 1B, 41, 0070C0...)
                    if color_hex.startswith("00") or color_hex.startswith("1B") or color_hex.startswith("41") or color_hex == "0070C0":
                        is_correct = True
                # Kiểm tra tô nền (Highlight)
                if run.font.highlight_color:
                    is_correct = True
            
            option_label = text[0] # Lấy ký tự A, B, C hoặc D
            current_q["options"].append(text)
            if is_correct:
                current_q["correct"] = option_label

    if current_q and len(current_q['options']) > 0:
        questions.append(current_q)
        
    return questions

# --- HÀM 2: XỬ LÝ FILE EXCEL (.XLSX) ---
def parse_xlsx(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb.active
    questions = []
    
    for row in sheet.iter_rows(min_row=2): # Bỏ qua dòng tiêu đề
        q_cell = row[1]   # Cột B: Nội dung câu hỏi
        ans_cell = row[2] # Cột C: Các phương án trả lời
        
        if not q_cell.value or not ans_cell.value:
            continue
            
        q_text = str(q_cell.value).strip()
        ans_text = str(ans_cell.value).strip()
        
        ans_lines = ans_text.split('\n')
        options = [line.strip() for line in ans_lines if line.strip()]
        
        correct_option = None
        
        # Kiểm tra màu chữ đỏ trong ô C (Cột phương án)
        if ans_cell.font and ans_cell.font.color and ans_cell.font.color.rgb:
            rgb_hex = str(ans_cell.font.color.rgb)
            # Kiểm tra mã màu đỏ (FF0000 hoặc ARGB có dải FF0000)
            if "FF0000" in rgb_hex or rgb_hex.startswith("FF00"):
                # Nhận diện phương án nào bị tô đỏ dựa theo cấu trúc A., B., C., D.
                for opt in options:
                    if opt.startswith(("A.", "B.", "C.", "D.", "A)", "B)", "C)", "D)")):
                        # Mặc định gán đáp án đúng tìm thấy
                        correct_option = opt[0]

        questions.append({
            "question": q_text,
            "options": options,
            "correct": correct_option or "A" # Nếu không phát hiện màu đỏ thì mặc định là A
        })
        
    return questions

# --- GIAO DIỆN UPLOAD FILE ---
uploaded_files = st.file_uploader(
    "Tải lên các file Word (.docx) hoặc Excel (.xlsx)", 
    type=["docx", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_qs = []
    for f in uploaded_files:
        if f.name.endswith(".docx"):
            qs = parse_docx(f)
            all_qs.extend(qs)
        elif f.name.endswith(".xlsx"):
            qs = parse_xlsx(f)
            all_qs.extend(qs)
            
    st.session_state["questions_bank"] = all_qs
    st.success(f"✅ Đã tải thành công {len(all_qs)} câu hỏi vào Ngân hàng dữ liệu!")

# --- TẠO ĐỀ THI NGẪU NHIÊN 50 CÂU ---
if st.session_state["questions_bank"]:
    total_available = len(st.session_state["questions_bank"])
    st.info(f"📊 Ngân hàng hiện tại có: **{total_available}** câu hỏi.")
    
    if st.button("🔀 Xáo trộn & Tạo đề thi 50 câu ngẫu nhiên"):
        sample_size = min(50, total_available)
        st.session_state["current_exam"] = random.sample(st.session_state["questions_bank"], sample_size)
        st.rerun()

# --- GIAO DIỆN LÀM BÀI VÀ NỘP BÀI ---
if st.session_state["current_exam"]:
    st.markdown("---")
    st.header("✍️ BÀI THI TRẮC NGHIỆM")
    
    user_answers = {}
    with st.form("exam_form"):
        for idx, q in enumerate(st.session_state["current_exam"], start=1):
            st.markdown(f"**Câu {idx}:** {q['question']}")
            
            # Chọn đáp án
            user_answers[idx] = st.radio(
                f"Chọn đáp án câu {idx}:", 
                q['options'], 
                key=f"q_{idx}",
                index=None
            )
            st.write("---")
            
        submitted = st.form_submit_button("📩 Nộp Bài Thi")
        
    if submitted:
        score = 0
        total_q = len(st.session_state["current_exam"])
        
        for idx, q in enumerate(st.session_state["current_exam"], start=1):
            selected = user_answers.get(idx)
            correct_key = q.get('correct')
            
            # So sánh đáp án người dùng chọn với đáp án đúng
            if selected and correct_key and selected.startswith(correct_key):
                score += 1
                
        st.balloons()
        st.header(f"🏆 Kết quả: {score}/{total_q} câu đúng ({(score/total_q)*100:.1f} điểm)")