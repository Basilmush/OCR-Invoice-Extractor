import streamlit as st
import pandas as pd
import numpy as np
import re
import pytesseract
import cv2
from pdf2image import convert_from_bytes
from openpyxl import Workbook
import io
from PIL import Image, ImageEnhance

# ================================
# ตั้งค่า Tesseract
# ================================
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

# ================================
# Helper Functions
# ================================
def preprocess_image_for_ocr(pil_img):
    enhancer = ImageEnhance.Contrast(pil_img)
    pil_img = enhancer.enhance(2.0)
    enhancer = ImageEnhance.Sharpness(pil_img)
    pil_img = enhancer.enhance(1.5)
    enhancer = ImageEnhance.Brightness(pil_img)
    pil_img = enhancer.enhance(1.1)

    img = np.array(pil_img.convert("RGB"))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 41, 15
    )
    denoised = cv2.medianBlur(thresh, 3)
    return denoised

def clean_amount(val):
    try:
        num = float(val.replace(",", ""))
        if num <= 0 or num > 50000:
            return ""
        return f"{num:.2f}"
    except:
        return ""

def extract_fields(text):
    data = {"date": "", "invoice_number": "", "amount": ""}

    # วันที่
    date_patterns = [r"(\d{1,2}/\d{1,2}/\d{2,4})", r"(\d{1,2}-\d{1,2}-\d{2,4})"]
    for p in date_patterns:
        m = re.search(p, text)
        if m:
            data["date"] = m.group(1)
            break

    # เลขที่บิล
    inv_patterns = [r"(HH\d{6,8})", r"(?:เลขที่|No\.?)\s*([A-Z0-9]{6,12})"]
    for p in inv_patterns:
        m = re.search(p, text)
        if m:
            data["invoice_number"] = m.group(1)
            break

    # ยอดก่อน VAT
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

# ================================
# Streamlit App
# ================================
st.title("📄 OCR Extractor for Multiple Pages (Editable Table)")
st.write("อัปโหลด PDF → OCR → แก้ไขค่าทั้งหมดในตารางเดียว → ดาวน์โหลด Excel")

uploaded_file = st.file_uploader("อัปโหลด PDF", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    with st.spinner("กำลังแปลง PDF เป็นภาพและ OCR ..."):
        pages = convert_from_bytes(pdf_bytes, dpi=400)
        results = []

        for i, page in enumerate(pages):
            proc = preprocess_image_for_ocr(page)
            text1 = pytesseract.image_to_string(proc, lang="tha+eng", config="--psm 6 --oem 3")
            text2 = pytesseract.image_to_string(proc, lang="tha+eng", config="--psm 11 --oem 3")
            text = text1 + "\n" + text2
            data = extract_fields(text)
            data["page_number"] = i + 1
            results.append(data)

    # สร้าง DataFrame สำหรับแก้ไข
    df = pd.DataFrame(results)
    df = df[["page_number", "date", "invoice_number", "amount"]]
    df.rename(columns={"page_number": "ลำดับ", "date": "วันที่",
                       "invoice_number": "เลขที่ตามบิล", "amount": "ยอดก่อน VAT"}, inplace=True)
    df["ลำดับ"] = df.index + 1

    st.subheader("📋 แก้ไขค่าทั้งหมดในตารางเดียว")
    edited_df = st.data_editor(df, num_rows="dynamic")  # Streamlit >= 1.24

    excel_file = fill_excel_with_data(edited_df[["วันที่", "เลขที่ตามบิล", "ยอดก่อน VAT"]].to_dict(orient="records"))
    st.download_button(
        "⬇️ ดาวน์โหลด Excel",
        data=excel_file,
        file_name="Invoice_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
