import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import Image
import os
import io

def process_pdf_data(pdf_bytes):
    status_placeholder = st.empty()
    try:
        with open("temp_upload.pdf", "wb") as f:
            f.write(pdf_bytes)

        status_placeholder.info("กำลังแปลง PDF เป็นรูปภาพ...")
        pages = convert_from_path("temp_upload.pdf", dpi=300)

        data = []
        for i, page in enumerate(pages):
            status_placeholder.info(f"กำลังประมวลผลหน้าที่ {i+1}/{len(pages)}...")
            
            text = pytesseract.image_to_string(page, lang="tha+eng")

            date_pattern = r"(\d{2}/\d{2}/\d{2})"
            no_pattern = r"HH\d{6,}"
            
            # Regex ขั้นสุดท้าย: ครอบคลุมคำว่า 'มูลค่าสินค้า' หรือ 'Product Value' 
            # และจับกลุ่มตัวเลขที่อยู่บรรทัดเดียวกันหรือบรรทัดถัดไป
            amount_pattern = r"(?:มูลค่าสินค้า|Product Value)\s*[\n\r]*\s*([,\d]+\.\d{2})"
            
            date_match = re.search(date_pattern, text)
            no_match = re.search(no_pattern, text)
            amount_match = re.search(amount_pattern, text, re.IGNORECASE)

            date_value = date_match.group(1) if date_match else "N/A"
            no_value = no_match.group(0) if no_match else "N/A"
            
            if amount_match:
                # group(1) คือตัวเลขที่ถูกจับกลุ่มใน Regex ใหม่
                amount_value = float(amount_match.group(1).replace(',', ''))
            else:
                amount_value = 0.0

            data.append([date_value, no_value, amount_value])
        
        os.remove("temp_upload.pdf")
        status_placeholder.empty()

        return pd.DataFrame(data, columns=['วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT'])

    except Exception as e:
        status_placeholder.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล: {e}")
        if os.path.exists("temp_upload.pdf"):
            os.remove("temp_upload.pdf")
        return pd.DataFrame()

st.set_page_config(page_title="PDF OCR Extractor", layout="wide")
st.title("เครื่องมือดึงข้อมูลใบเสร็จจาก PDF (OCR)")
st.markdown("---")

st.header("ขั้นตอนการใช้งานสำหรับเพื่อนของคุณ")
st.markdown("""
1. คลิกที่ปุ่ม **"Browse files"** ด้านล่างเพื่ออัปโหลดไฟล์ใบเสร็จ PDF
2. โปรแกรมจะทำการ **OCR** และดึงข้อมูล **วันที่**, **เลขที่ตามบิล**, และ **ยอดก่อน VAT**
3. ข้อมูลสรุปจะแสดงในตารางด้านล่าง และสามารถกดปุ่ม **"ดาวน์โหลดไฟล์ Excel** ได้ทันที
""")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ PDF ใบเสร็จรับเงิน", type="pdf")

if uploaded_file is not None:
    with st.spinner("กำลังเริ่มต้นประมวลผล..."):
        pdf_bytes = uploaded_file.getvalue()
        
        df = process_pdf_data(pdf_bytes)
    
    if not df.empty:
        st.success("✅ ประมวลผลเสร็จสิ้น! ข้อมูลทั้งหมดถูกดึงออกมาแล้ว")
        
        st.subheader("ข้อมูลสรุปที่ดึงออกมา")
        st.dataframe(df, use_container_width=True, height=300)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Summary Data')
        output.seek(0)
        
        st.download_button(
            label="⬇️ ดาวน์โหลดไฟล์ Excel (Summary Data)",
            data=output,
            file_name="Invoice_Extracted_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.error("ไม่สามารถดึงข้อมูลได้ โปรดลองตรวจสอบไฟล์ PDF หรือลองอัปโหลดอีกครั้ง")
