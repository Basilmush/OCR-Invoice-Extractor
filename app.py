import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import Image
import os
import io

def clean_and_extract_number(text):
    """ฟังก์ชันทำความสะอาดและแปลงตัวเลข"""
    if not text:
        return 0.0
    
    # ลบตัวอักษรและเครื่องหมายที่ไม่ใช่ตัวเลข จุด และคอมม่า
    cleaned = re.sub(r'[^\d.,]', '', text.strip())
    
    # จัดการกับรูปแบบตัวเลขที่มีคอมม่า
    if ',' in cleaned:
        cleaned = cleaned.replace(',', '')
    
    try:
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0

def extract_data_from_text(text, page_num):
    """ฟังก์ชันดึงข้อมูลจากข้อความที่ได้จาก OCR"""
    
    # Pattern สำหรับวันที่ - รองรับหลายรูปแบบ
    date_patterns = [
        r"(\d{1,2}/\d{1,2}/\d{2,4})",  # DD/MM/YY หรือ DD/MM/YYYY
        r"(\d{1,2}-\d{1,2}-\d{2,4})",  # DD-MM-YY หรือ DD-MM-YYYY
        r"(\d{2}\d{2}\d{2})"           # DDMMYY
    ]
    
    # Pattern สำหรับเลขที่เอกสาร - ครอบคลุมมากขึ้น
    invoice_patterns = [
        r"(HH\d{6,8})",               # HH + ตัวเลข 6-8 หลัก
        r"([A-Z]{2}\d{6,8})",         # ตัวอักษร 2 ตัว + ตัวเลข
        r"(INV[/-]?\d{6,8})",         # INV + ตัวเลข
        r"(\d{8,10})"                 # ตัวเลข 8-10 หลัก
    ]
    
    # Pattern สำหรับจำนวนเงิน - ครอบคลุมหลายรูปแบบ
    amount_patterns = [
        # รูปแบบที่มีคำว่า "มูลค่าสินค้า" หรือ "Product Value"
        r"(?:มูลค่าสินค้า|Product Value|รวม|Total|Amount)\s*:?\s*([,\d]+\.?\d*)",
        
        # รูปแบบตัวเลขที่มีจุดทศนิยม 2 หลัก
        r"([,\d]+\.\d{2})",
        
        # รูปแบบตัวเลขขนาดใหญ่ที่อาจเป็นจำนวนเงิน
        r"([,\d]{4,}\.?\d{0,2})",
        
        # รูปแบบที่มี VAT หรือ ก่อน VAT
        r"(?:ก่อน\s*VAT|Before\s*VAT|Pre\s*VAT)\s*:?\s*([,\d]+\.?\d*)",
        
        # รูปแบบบรรทัดที่มีตัวเลขใหญ่ๆ
        r"\s+([,\d]+\.?\d{2})\s*$",
    ]
    
    extracted_data = {
        'date': 'N/A',
        'invoice_no': 'N/A', 
        'amount': 0.0
    }
    
    # ค้นหาวันที่
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            date_str = match.group(1)
            # แปลงรูปแบบวันที่ให้เป็น DD/MM/YY
            if len(date_str) == 6:  # DDMMYY
                extracted_data['date'] = f"{date_str[:2]}/{date_str[2:4]}/{date_str[4:]}"
            else:
                extracted_data['date'] = date_str.replace('-', '/')
            break
    
    # ค้นหาเลขที่เอกสาร
    for pattern in invoice_patterns:
        matches = re.findall(pattern, text)
        if matches:
            # เลือกเลขที่ดูเหมือนเลขใบเสร็จมากที่สุด
            for match in matches:
                if match.startswith('HH') or len(match) >= 6:
                    extracted_data['invoice_no'] = match
                    break
            if extracted_data['invoice_no'] != 'N/A':
                break
    
    # ค้นหาจำนวนเงิน
    amounts = []
    for pattern in amount_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            amount = clean_and_extract_number(match)
            if amount > 0:
                amounts.append(amount)
    
    # เลือกจำนวนเงินที่เหมาะสมที่สุด
    if amounts:
        # เรียงลำดับและเลือกจำนวนที่ใหญ่ที่สุดแต่สมเหตุสมผล
        amounts = sorted(set(amounts), reverse=True)
        extracted_data['amount'] = amounts[0]
    
    return extracted_data

def process_pdf_data(pdf_bytes):
    status_placeholder = st.empty()
    try:
        # บันทึกไฟล์ชั่วคราว
        temp_file = "temp_upload.pdf"
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        
        status_placeholder.info("🔄 กำลังแปลง PDF เป็นรูปภาพ...")
        
        # แปลง PDF เป็นรูปภาพด้วย DPI สูงสำหรับความชัดเจน
        pages = convert_from_path(temp_file, dpi=400)
        
        data = []
        total_pages = len(pages)
        
        for i, page in enumerate(pages):
            status_placeholder.info(f"📖 กำลังประมวลผลหน้าที่ {i+1}/{total_pages}...")
            
            # ปรับปรุงคุณภาพรูปภาพก่อน OCR
            # เพิ่ม contrast และ sharpness
            from PIL import ImageEnhance
            enhancer = ImageEnhance.Contrast(page)
            page = enhancer.enhance(2.0)
            enhancer = ImageEnhance.Sharpness(page)
            page = enhancer.enhance(2.0)
            
            # ใช้ OCR กับรูปภาพที่ปรับปรุงแล้ว
            text = pytesseract.image_to_string(
                page, 
                lang="tha+eng",
                config='--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz/.-,:ก-๙ '
            )
            
            # ดึงข้อมูลจากข้อความ
            extracted = extract_data_from_text(text, i+1)
            
            data.append([
                extracted['date'],
                extracted['invoice_no'], 
                extracted['amount']
            ])
            
            # Debug: แสดงข้อความที่ดึงได้ (สำหรับการปรับปรุง)
            if st.session_state.get('debug_mode', False):
                st.text_area(f"Raw OCR Text - Page {i+1}", text[:500], key=f"debug_{i}")
        
        # ลบไฟล์ชั่วคราว
        os.remove(temp_file)
        status_placeholder.empty()
        
        # สร้าง DataFrame
        df = pd.DataFrame(data, columns=['วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT'])
        
        # กรองข้อมูลที่มีความน่าเชื่อถือ
        # เก็บเฉพาะแถวที่มีข้อมูลอย่างน้อย 2 จาก 3 คอลัมน์
        df_filtered = df[
            ((df['วันที่'] != 'N/A') + 
             (df['เลขที่ตามบิล'] != 'N/A') + 
             (df['ยอดก่อน VAT'] > 0)) >= 2
        ].copy()
        
        return df_filtered if not df_filtered.empty else df
        
    except Exception as e:
        status_placeholder.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return pd.DataFrame()

def main():
    st.set_page_config(
        page_title="PDF OCR Extractor - Enhanced", 
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("🔍 เครื่องมือดึงข้อมูลใบเสร็จจาก PDF (OCR) - เวอร์ชันปรับปรุง")
    st.markdown("---")
    
    # Sidebar สำหรับการตั้งค่า
    with st.sidebar:
        st.header("⚙️ การตั้งค่า")
        
        debug_mode = st.checkbox("โหมดดีบัก (แสดงข้อความ OCR)", key='debug_mode')
        
        st.markdown("### 📋 รูปแบบข้อมูลที่รองรับ")
        st.markdown("""
        **วันที่:** DD/MM/YY, DD-MM-YY, DDMMYY
        
        **เลขที่เอกสาร:** HH + ตัวเลข, INV + ตัวเลข
        
        **จำนวนเงิน:** รูปแบบทศนิยม 2 หลัก
        """)
    
    # คำแนะนำการใช้งาน
    with st.expander("📖 วิธีใช้งานและเคล็ดลับ", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **ขั้นตอนการใช้งาน:**
            1. 📁 อัปโหลดไฟล์ PDF
            2. ⏳ รอการประมวลผล OCR
            3. 📊 ตรวจสอบผลลัพธ์ในตาราง
            4. 💾 ดาวน์โหลดไฟล์ Excel
            """)
            
        with col2:
            st.markdown("""
            **เคล็ดลับสำหรับผลลัพธ์ที่ดี:**
            - ใช้ PDF ที่มีความชัดเจน
            - หลีกเลี่ยงไฟล์ที่สแกนด้วยความละเอียดต่ำ
            - ตรวจสอบว่าข้อความไม่เอียงหया buồng
            - ไฟล์ไม่ควรมีรูปภาพซ้อนทับข้อความ
            """)
    
    # ส่วนอัปโหลดไฟล์
    uploaded_file = st.file_uploader(
        "📄 อัปโหลดไฟล์ PDF ใบเสร็จรับเงิน", 
        type="pdf",
        help="เลือกไฟล์ PDF ที่ต้องการดึงข้อมูล"
    )
    
    if uploaded_file is not None:
        # แสดงข้อมูลไฟล์
        file_size = len(uploaded_file.getvalue()) / 1024 / 1024  # MB
        st.info(f"📊 ไฟล์: {uploaded_file.name} | ขนาด: {file_size:.2f} MB")
        
        # ประมวลผลไฟล์
        with st.spinner("⚡ กำลังเริ่มต้นประมวลผล..."):
            pdf_bytes = uploaded_file.getvalue()
            df = process_pdf_data(pdf_bytes)
        
        if not df.empty:
            st.success(f"✅ ประมวลผลเสร็จสิ้น! พบข้อมูล {len(df)} รายการ")
            
            # แสดงสถิติสรุป
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("จำนวนหน้าที่ประมวลผล", len(df))
            with col2:
                valid_amounts = df[df['ยอดก่อน VAT'] > 0]['ยอดก่อน VAT'].count()
                st.metric("รายการที่มีจำนวนเงิน", valid_amounts)
            with col3:
                total_amount = df['ยอดก่อน VAT'].sum()
                st.metric("ยอดรวม", f"{total_amount:,.2f}")
            
            st.markdown("---")
            
            # แสดงตารางข้อมูล
            st.subheader("📋 ข้อมูลที่ดึงออกมา")
            
            # เพิ่มการแสดงผลแบบมีสี
            def highlight_rows(row):
                if row['ยอดก่อน VAT'] > 0:
                    return ['background-color: #d4edda'] * len(row)
                else:
                    return ['background-color: #f8d7da'] * len(row)
            
            styled_df = df.style.apply(highlight_rows, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=400)
            
            # ส่วนดาวน์โหลด
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            
            with col2:
                # สร้างไฟล์ Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Invoice_Data')
                    
                    # เพิ่มสถิติในชีทใหม่
                    summary_data = {
                        'รายการ': ['จำนวนหน้าทั้งหมด', 'รายการที่มีจำนวนเงิน', 'ยอดรวมทั้งหมด'],
                        'ค่า': [len(df), valid_amounts, total_amount]
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, index=False, sheet_name='Summary')
                
                output.seek(0)
                
                st.download_button(
                    label="⬇️ ดาวน์โหลดไฟล์ Excel",
                    data=output,
                    file_name=f"Invoice_Data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            
        else:
            st.error("❌ ไม่สามารถดึงข้อมูลได้ กรุณาตรวจสอบไฟล์ PDF และลองใหม่อีกครั้ง")
            st.markdown("""
            **คำแนะนำ:**
            - ตรวจสอบว่าไฟล์ PDF ไม่เสียหาย
            - ลองใช้ไฟล์ที่มีความชัดเจนมากกว่านี้
            - ตรวจสอบว่าข้อความในไฟล์เป็นภาษาไทยหรือภาษาอังกฤษ
            """)

if __name__ == "__main__":
    main()
