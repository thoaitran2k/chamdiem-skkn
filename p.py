import docx
import os
import re

def clean_text(text):
    # Loại bỏ các ký tự xuống dòng dư thừa và khoảng trắng
    return " ".join(text.split())

def extract_info_from_doc(file_path):
    doc = docx.Document(file_path)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    
    # Sử dụng Regex để tìm tên đề tài - bất kể nó có ngoặc kép hay không
    # Tìm đoạn văn bản sau chữ "công nhận sáng kiến:"
    match = re.search(r"công nhận sáng kiến:\s*[""“]?([^”""]+)", full_text, re.IGNORECASE)
    ten_de_tai = match.group(1).strip() if match else "Không tìm thấy"
    
    return {
        "ten_de_tai": ten_de_tai,
        # Bạn có thể thêm các mỏ neo khác tương tự cho Tên tác giả, Đơn vị...
    }

# Chạy cho cả thư mục
folder_path = "D:/THCSVT/he-thong-cham-diem-skkn/SKKN-XA-LAP-VO"
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        data = extract_info_from_doc(os.path.join(folder_path, filename))
        print(f"File: {filename} | Tên đề tài: {data['ten_de_tai']}")