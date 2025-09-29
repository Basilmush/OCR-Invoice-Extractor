import streamlit as st
import pandas as pd
import numpy as np
import pytesseract
import cv2
from pdf2image import convert_from_bytes
from PIL import Image, ImageEnhance
import io
import re
from concurrent.futures import ThreadPoolExecutor


pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"


def preprocess_and_ocr(pil_img):
    """ปรับภาพ + OCR แบบ optimized"""
   
    w, h = pil_img.size
    max_width = 1800
    if w > max_width:
        ratio = max_width / w
        pil_img = pil_img.resize((max_width, int(h*ratio)))

  
    enhancer = ImageEnhance.Contrast(pil_img)
    pil_img = enhancer.enhance(2.0)
    enhancer = ImageEnhance.Sharpness(pil_img)
    pil_img = enhancer.enhance(1.5)
    enhancer = ImageEnhance.Brightness(pil_img)
    pil_img = enhancer.enhance(1.1)

  
    img = np.array(pil_img.convert("L"))
    thresh = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 41, 15)
    denoised = cv2.medianBlur(thresh, 3)

   
    text = pytesseract.image_to_string(denoised, lang="tha+eng", config="--psm 6 --oem 3")
    return text

def clean_amount(val):
    try:
        num = float(val.replace(",", ""))
        if num <= 0 or num > 50000:
            return ""
        return f"{num:.2f}"
    except:
        return ""

def extract_fields(text):
    """ดึงวันที่, เลขที่บิล, ยอดก่อน VAT"""
    data = {"date": "", "invoice_number": "", "amount": ""}

    date_patterns = [r"(\d{1,2}/\d{1,2}/\d{2,4})", r"(\d{1,2}-\d{1,2}-\d{2,4})"]
    for p in date_patterns:
        m = re.search(p, text)
        if m:
            data["date"] = m.group(1)
            break

    inv_patterns = [r"(HH\d{6,8})", r"(?:เลขที่|No\.?)\s*([A-Z0-9]{6,12})"]
    for p in inv_patterns:
        m = re.search(p, text)
        if m:
            data["invoice_number"] = m.group(1)
            break

    amt_patterns = [
        r"(?:มูลค่าสินค้า|Subtotal|ก่อนภาษี|จำนวนเงินรวมทั้งสิ้น|Total).*?([0-9,]+\.\d{2})",
        r"([0-9,]+\.\d{2})"
    ]
    for p in amt_patterns:
        m = re.search(p, text, re.DOTALL | re.IGNORECASE)
        if m:
            data["amount"] = clean_amount(m.group(1))
            if data["amount"]:
                break

    return data

def fill_excel_with_data(data_list):
    df = pd.DataFrame(data_list)
    df.insert(0, "ลำดับ", df.index + 1)
    df.columns = ["ลำดับ", "วันที่", "เลขที่ตามบิล", "ยอดก่อน VAT"]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Invoice_Data")
    output.seek(0)
    return output


st.set_page_config(page_title="OCR Invoice Tool", layout="wide")
st.title("📄 OCR & Editable Invoice Viewer")
st.write("อัปโหลด PDF / รูปภาพ → OCR → แก้ไขใต้ภาพ → ดาวน์โหลด Excel")

uploaded_file = st.file_uploader("อัปโหลด PDF หรือ รูปภาพ", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    if 'results' not in st.session_state:
        st.session_state.results = []

    if not st.session_state.results:
        with st.spinner("กำลังประมวลผล OCR ..."):
         
            if uploaded_file.type == "application/pdf":
                pages = convert_from_bytes(uploaded_file.read(), dpi=300)
            else:
                pages = [Image.open(uploaded_file).convert("RGB")]

            def process_page(page, idx):
                text = preprocess_and_ocr(page)
                data = extract_fields(text)
                data["page_number"] = idx + 1
                data["image"] = page
                return data

            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(process_page, p, i) for i, p in enumerate(pages)]
                for f in futures:
                    st.session_state.results.append(f.result())

    st.subheader("📋 แก้ไขข้อมูลแต่ละหน้า")
    edited_results = []
    for row in st.session_state.results:
        st.markdown(f"### หน้า {row['page_number']}")
        st.image(row["image"], use_column_width=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            date_val = st.text_input(f"วันที่ หน้า {row['page_number']}", value=row["date"], key=f"date_{row['page_number']}")
        with col2:
            inv_val = st.text_input(f"เลขที่บิล หน้า {row['page_number']}", value=row["invoice_number"], key=f"inv_{row['page_number']}")
        with col3:
            amt_val = st.text_input(f"ยอดก่อน VAT หน้า {row['page_number']}", value=row["amount"], key=f"amt_{row['page_number']}")

        edited_results.append({"date": date_val, "invoice_number": inv_val, "amount": amt_val})

    st.subheader("💾 ดาวน์โหลด Excel")
    excel_file = fill_excel_with_data(edited_results)
    st.download_button(
        "⬇️ ดาวน์โหลด Excel",
        data=excel_file,
        file_name="Invoice_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
