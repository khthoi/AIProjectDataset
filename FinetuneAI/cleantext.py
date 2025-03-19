import fitz  # PyMuPDF
import openpyxl
import re

def remove_illegal_chars(text):
    """Loại bỏ các ký tự không hợp lệ trong văn bản để ghi vào Excel"""
    if text is None:
        return ""
    return re.sub(r'[\x00-\x08\x0B-\x1F\x7F]', '', text)  # Xóa ký tự điều khiển

def split_text_into_paragraphs(text, min_words=350, max_words=750):
    """Chia văn bản thành các đoạn liền mạch dựa vào khoảng trắng lớn hoặc dấu câu"""
    paragraphs = re.split(r'\n\n+|\r\n\r\n+', text)  # Tách theo đoạn văn thực tế
    meaningful_paragraphs = []

    for para in paragraphs:
        words = para.split()
        if len(words) < min_words:
            continue  # Bỏ qua đoạn quá ngắn

        if len(words) > max_words:
            # Chia đoạn dài thành nhiều đoạn nhỏ nhưng vẫn liền mạch
            for i in range(0, len(words), max_words):
                sub_para = " ".join(words[i:i+max_words])
                meaningful_paragraphs.append(sub_para)
        else:
            meaningful_paragraphs.append(para.strip())

    return meaningful_paragraphs

def extract_text_from_pdf(pdf_path, start_page=0, min_words=350, max_words=750, num_samples=90):
    doc = fitz.open(pdf_path)
    extracted_texts = []

    for page_num in range(start_page, len(doc)):  # Chỉ duyệt từ start_page trở đi
        page = doc[page_num]
        text = page.get_text("text").strip()

        if len(text.split()) < min_words:
            continue  # Bỏ qua trang có quá ít từ
        
        # Chia thành các đoạn liền mạch
        paragraphs = split_text_into_paragraphs(text, min_words, max_words)

        for paragraph in paragraphs:
            extracted_texts.append(paragraph)
            if len(extracted_texts) >= num_samples:
                return extracted_texts  # Dừng khi đủ số lượng

    return extracted_texts

def save_to_excel(text_list, excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Nội dung trích xuất"])

    for text in text_list:
        clean_text = remove_illegal_chars(text)  # Làm sạch văn bản trước khi lưu
        ws.append([clean_text])

    wb.save(excel_path)
    print(f"Đã lưu vào {excel_path}")

# --- Chạy chương trình ---
pdf_file = "test.pdf"  # Đường dẫn file PDF đầu vào
excel_file = "output.xlsx"
start_page = 1  # Chỉ bắt đầu từ trang 712

text_data = extract_text_from_pdf(pdf_file, start_page=start_page, num_samples=128)
if text_data:
    save_to_excel(text_data, excel_file)
