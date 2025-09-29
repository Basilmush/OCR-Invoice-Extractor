import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import Image, ImageEnhance, ImageFilter
import os
import io
import traceback
from datetime import datetime, timedelta
from streamlit.logger import get_logger

logger = get_logger(__name__)

# =========================================================
# การตั้งค่า Tesseract Path สำหรับ Cloud Server
# =========================================================
try:
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
except Exception as e:
    st.warning(f"⚠️ การตั้งค่า Tesseract ล้มเหลว: {e}. ตรวจสอบการติดตั้ง Tesseract.")
    logger.warning(f"Tesseract setup failed: {e}")

def custom_exception_handler(exc_type, exc_value, exc_traceback):
    st.error(f"❌ เกิดข้อผิดพลาดไม่คาดคิด: {exc_value}")
    logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    st.session_state.error_log = traceback.format_exc()
    if st.button("📋 ดูรายละเอียดข้อผิดพลาด"):
        st.text(st.session_state.error_log)

# Override global exception handler
import sys
sys.excepthook = custom_exception_handler

def enhance_image_for_ocr(image):
    """ปรับปรุงคุณภาพรูปภาพสำหรับ OCR"""
    try:
        # แปลงเป็น grayscale
        image = image.convert('L')
        # เพิ่ม contrast
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.5)
        # เพิ่ม sharpness
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        # กรองภาพเพื่อลด noise
        image = image.filter(ImageFilter.UnsharpMask(radius=2, percent=200, threshold=3))
        # แปลงเป็น binary (ขาว-ดำ)
        image = image.point(lambda x: 0 if x < 140 else 255, '1')
        return image
    except Exception as e:
        st.warning(f"⚠️ การปรับภาพล้มเหลว: {e}. ใช้ภาพดั้งเดิมแทน.")
        logger.warning(f"Image enhancement failed: {e}")
        return image

def extract_ocr_from_pdf(pdf_bytes):
    """แปลง PDF เป็น OCR Text และคืนค่าทั้ง text และ images"""
    temp_file = "temp_upload.pdf"
    try:
        if not pdf_bytes or len(pdf_bytes) == 0:
            raise ValueError("ไฟล์ PDF ว่างเปล่า")
        # บันทึกไฟล์ชั่วคราว
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        
        st.info("🔄 กำลังแปลง PDF เป็นรูปภาพ...")
        pages = convert_from_path(temp_file, dpi=600, fmt='PNG')  # เพิ่ม DPI เป็น 600
        
        ocr_results = []
        for i, page in enumerate(pages):
            st.info(f"📖 กำลังทำ OCR หน้าที่ {i+1}/{len(pages)}...")
            enhanced_page = enhance_image_for_ocr(page)
            ocr_text = pytesseract.image_to_string(
                enhanced_page,
                lang="tha+eng",
                config='--psm 6 --oem 3'
            )
            ocr_results.append({
                'page_number': i + 1,
                'ocr_text': ocr_text,
                'image': enhanced_page
            })
        
        os.remove(temp_file)
        return ocr_results
    except Exception as e:
        if os.path.exists(temp_file):
            os.remove(temp_file)
        st.error(f"❌ เกิดข้อผิดพลาดในการแปลง PDF: {str(e)}")
        logger.error(f"PDF conversion failed: {e}")
        return []

def clean_amount(raw_amount):
    """ทำความสะอาดตัวเลขที่ดึงมา (ลบคอมมาและแปลงเป็นทศนิยมสองหลัก)"""
    if not raw_amount:
        return ""
    cleaned = re.sub(r'[^\d\.]', '', raw_amount.replace(',', ''))
    try:
        return f"{float(cleaned):.2f}"
    except ValueError:
        return ""

def extract_data_from_ocr_text(text):
    """ดึงข้อมูลจากข้อความ OCR พร้อมคะแนนความน่าจะเป็น"""
    data = {'date': '', 'invoice_number': '', 'amount': '', 'raw_matches': {}, 'confidence': 0}
    
    # วันที่
    date_pattern = r'(?:วันที|Date|วันที่)\s*[.,:\s\n\r]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})'
    date_matches = re.search(date_pattern, text, re.IGNORECASE)
    if date_matches:
        data['date'] = date_matches.group(1)
        data['confidence'] += 30
        data['raw_matches']['dates_found'] = [data['date']]
    
    # เลขที่
    invoice_pattern = r'(?:เลขที|No|เลขที่)\s*[.,:\s\n\r]*\s*([H][H]\d{6,8}|\w{2}\d{6,8})'
    invoice_matches = re.search(invoice_pattern, text, re.IGNORECASE)
    if invoice_matches:
        data['invoice_number'] = invoice_matches.group(1)
        data['confidence'] += 30
        data['raw_matches']['invoices_found'] = [data['invoice_number']]
    
    # ยอดเงิน
    fuzzy_pattern = r"(?:[มม]*ูลค่าสินค้า|Product\s*Value|Net\s*Amount)\s*[.,:\s\n\r]*\s*([,\d]+\.\d{2})\s*(?:บาท)?"
    deep_fallback_pattern = r"(?:หักส่วนลด|Less\s*Discount)(?:.|\n)*?([,\d]+\.\d{2})\s*(?:จำนวนภาษีมูลค่าเพิ่ม|7\.00\s*%)"
    amount_match = re.search(fuzzy_pattern, text, re.IGNORECASE | re.DOTALL)
    if not amount_match:
        amount_match = re.search(deep_fallback_pattern, text, re.IGNORECASE | re.DOTALL)
    if amount_match:
        raw_amount = amount_match.group(1)
        data['amount'] = clean_amount(raw_amount)
        data['confidence'] += 40
        data['raw_matches']['amounts_found'] = [data['amount']]
    
    return data

def validate_data(data):
    """ตรวจสอบความสมเหตุสมผลของข้อมูล"""
    try:
        if data['date']:
            date_obj = datetime.strptime(data['date'], '%d/%m/%Y')
            if date_obj > datetime.now() or date_obj < datetime.now() - timedelta(days=365*2):
                data['date'] = ''
                st.warning(f"⚠️ วันที่ {data['date']} ไม่สมเหตุสมผล กรุณาตรวจสอบ")
        if data['amount']:
            amount = float(data['amount'].replace(',', '')) if data['amount'].replace('.', '').isdigit() else 0
            if amount > 500000 or amount < 0:
                data['amount'] = ''
                st.warning(f"⚠️ ยอดเงิน {data['amount']} ไม่สมเหตุสมผล กรุณาตรวจสอบ")
        return data
    except ValueError:
        return data

def create_excel_template():
    """สร้างไฟล์ Excel Template"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice_Data"
    headers = ['ลำดับ', 'วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT']
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def fill_excel_with_data(data_list):
    """กรอกข้อมูลลง Excel"""
    df_data = pd.DataFrame(data_list)
    df_data = df_data.sort_values(by='page_number').reset_index(drop=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_excel = df_data[['date', 'invoice_number', 'amount']].copy()
        df_to_excel.insert(0, 'ลำดับ', df_to_excel.index + 1)
        df_to_excel.columns = ['ลำดับ', 'วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT']
        df_to_excel.to_excel(writer, index=False, sheet_name='Invoice_Data')
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="PDF OCR Extractor", layout="wide", initial_sidebar_state="expanded")
    st.title("🔍 PDF OCR Checker & Excel Data Filler")
    st.markdown("**(สำหรับใบเสร็จ บริษัท ธนารัตน์ปิยะปิโตรเลียม จำกัด)**")
    st.markdown("---")
    
    if 'ocr_results' not in st.session_state:
        st.session_state.ocr_results = []
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = []
    
    with st.sidebar:
        st.header("⚙️ สรุปขั้นตอน")
        st.markdown("""
        1. 📁 **อัปโหลด PDF**
        2. 🚀 **กด 'เริ่ม OCR'**
        3. ✏️ **ตรวจสอบ/แก้ไข (สำคัญ!)**
        4. 💾 **ดาวน์โหลด Excel**
        """)
        st.markdown("---")
        st.markdown("### 🎯 ข้อมูลที่ต้องการ:")
        st.markdown("""
        - **วันที่** (รูปแบบ: DD/MM/YYYY)
        - **เลขที่ตามบิล** (เช่น HH6800470)
        - **มูลค่าสินค้า** (เช่น 4710.28)
        """)
    
    st.header("1. 📁 อัปโหลดไฟล์ PDF")
    uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf", help="อัปโหลดไฟล์ใบเสร็จ PDF ที่รวมหลายหน้าได้")
    
    if uploaded_file is not None:
        st.success(f"✅ อัปโหลดไฟล์: {uploaded_file.name}")
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("🚀 เริ่มแปลง OCR", type="primary"):
                with st.spinner("⏳ กำลังประมวลผล PDF และทำ OCR..."):
                    pdf_bytes = uploaded_file.getvalue()
                    st.session_state.ocr_results = extract_ocr_from_pdf(pdf_bytes)
        
        if st.session_state.ocr_results:
            st.success(f"✅ แปลง OCR เสร็จสิ้น! พบ {len(st.session_state.ocr_results)} หน้า")
            st.header("2. 👁️ ตรวจสอบและแก้ไขข้อมูล (จำเป็น)")
            st.warning("⚠️ โปรดตรวจสอบข้อมูลที่ดึงมาโดย OCR และกด 'บันทึก' ทุกหน้า")
            
            for result in st.session_state.ocr_results:
                page_key = result['page_number']
                extracted = extract_data_from_ocr_text(result['ocr_text'])
                extracted = validate_data(extracted)  # ตรวจสอบข้อมูล
                saved_data = next((d for d in st.session_state.extracted_data if d['page_number'] == page_key), None)
                
                with st.expander(f"📄 ใบเสร็จหน้าที่ {page_key} (ความแม่นยำ: {extracted['confidence']}%)", expanded=False):
                    st.markdown(f"**สถานะ:** {'💾 บันทึกแล้ว' if saved_data else '✏️ รอการบันทึก/แก้ไข'}")
                    if extracted['confidence'] < 80:
                        st.warning("⚠️ ความแม่นยำต่ำ กรุณาตรวจสอบข้อมูลอย่างละเอียด")
                    
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        st.subheader("📝 ข้อมูลที่ดึงได้:")
                        initial_date = saved_data['date'] if saved_data else extracted['date']
                        initial_invoice = saved_data['invoice_number'] if saved_data else extracted['invoice_number']
                        initial_amount = saved_data['amount'] if saved_data else extracted['amount']
                        
                        date_value = st.text_input("📅 วันที่:", value=initial_date, key=f"date_{page_key}")
                        invoice_value = st.text_input("🔢 เลขที่ตามบิล:", value=initial_invoice, key=f"invoice_{page_key}")
                        amount_value = st.text_input("💰 ยอดก่อน VAT:", value=initial_amount, key=f"amount_{page_key}")
                        
                        if st.button(f"💾 บันทึก/อัปเดตข้อมูลหน้า {page_key}", key=f"save_{page_key}", type="primary", use_container_width=True):
                            existing_index = next((i for i, data in enumerate(st.session_state.extracted_data) if data['page_number'] == page_key), None)
                            page_data = {'page_number': page_key, 'date': date_value, 'invoice_number': invoice_value, 'amount': amount_value}
                            if existing_index is not None:
                                st.session_state.extracted_data[existing_index] = page_data
                            else:
                                st.session_state.extracted_data.append(page_data)
                            st.rerun()
                    
                    with col2:
                        st.subheader("🖼์ รูปภาพ OCR & Text (ต้นฉบับ):")
                        st.image(result['image'], use_container_width=True)
                        with st.expander("📝 ข้อความ OCR ดิบ (Raw Text)"):
                            st.text_area("OCR Text:", result['ocr_text'], height=250, key=f"ocr_text_{page_key}")
            
            if st.session_state.extracted_data:
                st.header("3. 📊 แก้ไขข้อมูลรวม (ถ้าจำเป็น)")
                st.info("คุณสามารถแก้ไขข้อมูลทั้งหมดในตารางนี้ก่อนดาวน์โหลด")
                global_data = [{'ลำดับ': i+1, 'วันที่': d['date'], 'เลขที่ตามบิล': d['invoice_number'], 'ยอดก่อน VAT': d['amount']} 
                              for i, d in enumerate(sorted(st.session_state.extracted_data, key=lambda x: x['page_number']))]
                global_df = pd.DataFrame(global_data)
                search_term = st.text_input("🔍 ค้นหาในตารางรวม", placeholder="พิมพ์เพื่อค้นหา...")
                if search_term:
                    global_df = global_df[global_df.apply(lambda row: search_term.lower() in str(row).lower(), axis=1)]
                edited_global_df = st.data_editor(
                    global_df,
                    column_config={
                        "ลำดับ": st.column_config.NumberColumn("ลำดับ", disabled=True),
                        "วันที่": st.column_config.TextColumn("วันที่", help="รูปแบบ: DD/MM/YYYY"),
                        "เลขที่ตามบิล": st.column_config.TextColumn("เลขที่ตามบิล", help="เช่น HH6800470"),
                        "ยอดก่อน VAT": st.column_config.TextColumn("ยอดก่อน VAT", help="ตัวอย่าง: 4710.28")
                    },
                    use_container_width=True,
                    num_rows="dynamic"
                )
                if st.button("💾 บันทึกการเปลี่ยนแปลงรวม", type="primary"):
                    for i, row in edited_global_df.iterrows():
                        page_key = i + 1
                        existing = next((d for d in st.session_state.extracted_data if d['page_number'] == page_key), None)
                        if existing:
                            existing['date'] = row['วันที่']
                            existing['invoice_number'] = row['เลขที่ตามบิล']
                            existing['amount'] = row['ยอดก่อน VAT']
                    st.success("✅ อัปเดตข้อมูลรวมสำเร็จ!")
                    st.rerun()
            
            if st.session_state.extracted_data:
                st.header("4. 💾 ดาวน์โหลดไฟล์ Excel")
                df_summary = pd.DataFrame(sorted(st.session_state.extracted_data, key=lambda x: x['page_number']))
                df_summary = df_summary.reset_index(drop=True)
                st.subheader("📋 สรุปข้อมูลที่ถูกบันทึกแล้ว:")
                st.dataframe(df_summary, use_container_width=True, height=300)
                st.info(f"📊 พร้อมสร้างไฟล์ Excel ด้วยข้อมูล {len(st.session_state.extracted_data)} รายการ")
                excel_file = fill_excel_with_data(st.session_state.extracted_data)
                st.download_button(
                    label="⬇️ ดาวน์โหลด Excel (Final File)",
                    data=excel_file,
                    file_name=f"Invoice_Data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        
    else:
        st.info("👆 กรุณาอัปโหลดไฟล์ PDF เพื่อเริ่มต้น")
        st.header("📋 หรือดาวน์โหลด Excel Template")
        template_file = create_excel_template()
        st.download_button(
            label="⬇️ ดาวน์โหลด Excel Template",
            data=template_file,
            file_name="Invoice_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
