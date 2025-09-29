import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageEnhance
import os
import io

def enhance_image_for_ocr(image):
    """ปรับปรุงคุณภาพรูปภาพสำหรับ OCR"""
    # เพิ่ม contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    
    # เพิ่ม sharpness
    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(1.5)
    
    # เพิ่ม brightness เล็กน้อย
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(1.1)
    
    return image

def extract_ocr_from_pdf(pdf_bytes):
    """แปลง PDF เป็น OCR Text และคืนค่าทั้ง text และ images"""
    temp_file = "temp_upload.pdf"
    try:
        # บันทึกไฟล์ชั่วคราว
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        
        st.info("🔄 กำลังแปลง PDF เป็นรูปภาพ...")
        
        # แปลง PDF เป็นรูปภาพ
        pages = convert_from_path(temp_file, dpi=400)
        
        ocr_results = []
        processed_images = []
        
        for i, page in enumerate(pages):
            st.info(f"📖 กำลังทำ OCR หน้าที่ {i+1}/{len(pages)}...")
            
            # ปรับปรุงคุณภาพรูปภาพ
            enhanced_page = enhance_image_for_ocr(page)
            processed_images.append(enhanced_page)
            
            # ทำ OCR
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
        
        # ลบไฟล์ชั่วคราว
        os.remove(temp_file)
        
        return ocr_results
        
    except Exception as e:
        if os.path.exists(temp_file):
            os.remove(temp_file)
        st.error(f"❌ เกิดข้อผิดพลาดในการแปลง PDF: {str(e)}")
        return []

def extract_data_from_ocr_text(text):
    """ดึงข้อมูลจากข้อความ OCR"""
    data = {
        'date': '',
        'invoice_number': '',
        'amount': '',
        'raw_matches': {}
    }
    
    # Pattern สำหรับวันที่ - รูปแบบ DD/MM/YY
    date_pattern = r'(\d{1,2}/\d{1,2}/\d{2,4})'
    date_matches = re.findall(date_pattern, text)
    if date_matches:
        data['date'] = date_matches[0]
        data['raw_matches']['dates_found'] = date_matches
    
    # Pattern สำหรับเลข HH
    invoice_pattern = r'(HH\d{6,8})'
    invoice_matches = re.findall(invoice_pattern, text)
    if invoice_matches:
        data['invoice_number'] = invoice_matches[0]
        data['raw_matches']['invoices_found'] = invoice_matches
    
    # Pattern สำหรับมูลค่าสินค้า - หลายรูปแบบ
    amount_patterns = [
        r'มูลค่าสินค้า\s*:?\s*([,\d]+\.?\d{0,2})',  # มูลค่าสินค้า: 4710.28
        r'Product\s*Value\s*:?\s*([,\d]+\.?\d{0,2})',  # Product Value: 4710.28
        r'(\d{1,3}(?:,\d{3})*\.\d{2})',  # รูปแบบตัวเลข 4,710.28 หรือ 4710.28
    ]
    
    all_amounts = []
    for pattern in amount_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        all_amounts.extend(matches)
    
    if all_amounts:
        # เลือกตัวเลขที่มีทศนิยม 2 หลักและมีขนาดเหมาะสม
        valid_amounts = []
        for amount in all_amounts:
            clean_amount = amount.replace(',', '')
            try:
                float_amount = float(clean_amount)
                if 100 <= float_amount <= 1000000:  # ช่วงที่เหมาะสม
                    valid_amounts.append(amount)
            except ValueError:
                continue
        
        if valid_amounts:
            data['amount'] = valid_amounts[0]
        data['raw_matches']['amounts_found'] = all_amounts
    
    return data

def create_excel_template():
    """สร้างไฟล์ Excel Template"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice_Data"
    
    # Header
    headers = ['ลำดับ', 'วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT']
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # สร้างไฟล์ในหน่วยความจำ
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def fill_excel_with_data(data_list):
    """กรอกข้อมูลลง Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice_Data"
    
    # Header
    headers = ['ลำดับ', 'วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT']
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # กรอกข้อมูล
    for row_idx, data in enumerate(data_list, start=2):
        ws.cell(row=row_idx, column=1, value=row_idx-1)  # ลำดับ
        ws.cell(row=row_idx, column=2, value=data['date'])  # วันที่
        ws.cell(row=row_idx, column=3, value=data['invoice_number'])  # เลขที่ตามบิล
        ws.cell(row=row_idx, column=4, value=data['amount'])  # ยอดก่อน VAT
    
    # บันทึกลงหน่วยความจำ
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="PDF OCR Checker & Excel Filler", 
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("🔍 PDF OCR Checker & Excel Data Filler")
    st.markdown("**ตรวจสอบ OCR ก่อนกรอกข้อมูลลง Excel**")
    st.markdown("---")
    
    # Initialize session state
    if 'ocr_results' not in st.session_state:
        st.session_state.ocr_results = []
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = []
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ ขั้นตอนการทำงาน")
        st.markdown("""
        1. 📁 **อัปโหลด PDF**
        2. 👁️ **ตรวจสอบ OCR**
        3. ✏️ **แก้ไขข้อมูล (ถ้าจำเป็น)**
        4. 💾 **ดาวน์โหลด Excel**
        """)
        
        st.markdown("---")
        st.markdown("### 🎯 รูปแบบข้อมูลที่ต้องการ:")
        st.markdown("""
        - **วันที่:** 01/08/68
        - **เลขที่:** HH6800470  
        - **มูลค่า:** 4710.28
        """)
    
    # Step 1: Upload PDF
    st.header("📁 Step 1: อัปโหลดไฟล์ PDF")
    uploaded_file = st.file_uploader(
        "เลือกไฟล์ PDF", 
        type="pdf",
        help="อัปโหลดไฟล์ PDF ที่ต้องการแปลงเป็น OCR"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ อัปโหลดไฟล์: {uploaded_file.name}")
        
        # Step 2: Convert to OCR
        st.header("🔄 Step 2: แปลง PDF เป็น OCR")
        
        col1, col2 = st.columns([1, 4])
        with col1:
            if st.button("🚀 เริ่มแปลง OCR", type="primary"):
                with st.spinner("กำลังประมวลผล..."):
                    pdf_bytes = uploaded_file.getvalue()
                    st.session_state.ocr_results = extract_ocr_from_pdf(pdf_bytes)
                
                if st.session_state.ocr_results:
                    st.success(f"✅ แปลง OCR เสร็จสิ้น! จำนวนหน้า: {len(st.session_state.ocr_results)}")
        
        # Step 3: Display OCR Results
        if st.session_state.ocr_results:
            st.header("👁️ Step 3: ตรวจสอบผลลัพธ์ OCR")
            
            # แสดงผลลัพธ์แต่ละหน้า
            for result in st.session_state.ocr_results:
                with st.expander(f"📄 หน้าที่ {result['page_number']}", expanded=True):
                    
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.subheader("🖼️ รูปภาพที่ปรับปรุงแล้ว:")
                        st.image(result['image'], use_container_width=True)
                    
                    with col2:
                        st.subheader("📝 ข้อความที่ได้จาก OCR:")
                        st.text_area(
                            "OCR Text:", 
                            result['ocr_text'], 
                            height=300,
                            key=f"ocr_text_{result['page_number']}"
                        )
                        
                        # ดึงข้อมูลจาก OCR
                        extracted = extract_data_from_ocr_text(result['ocr_text'])
                        
                        st.subheader("🎯 ข้อมูลที่ดึงได้:")
                        
                        # แสดงข้อมูลที่ดึงได้
                        date_value = st.text_input(
                            "📅 วันที่:", 
                            value=extracted['date'],
                            key=f"date_{result['page_number']}"
                        )
                        
                        invoice_value = st.text_input(
                            "🔢 เลขที่ตามบิล:", 
                            value=extracted['invoice_number'],
                            key=f"invoice_{result['page_number']}"
                        )
                        
                        amount_value = st.text_input(
                            "💰 ยอดก่อน VAT:", 
                            value=extracted['amount'],
                            key=f"amount_{result['page_number']}"
                        )
                        
                        # แสดงข้อมูลที่พบทั้งหมด (สำหรับ Debug)
                        if extracted['raw_matches']:
                            with st.expander("🔍 ข้อมูลทั้งหมดที่พบ (Debug)"):
                                st.json(extracted['raw_matches'])
                        
                        # บันทึกข้อมูลที่แก้ไขแล้ว
                        if st.button(f"💾 บันทึกข้อมูลหน้า {result['page_number']}", key=f"save_{result['page_number']}"):
                            # ตรวจสอบว่ามีข้อมูลหน้านี้อยู่แล้วหรือไม่
                            existing_index = None
                            for i, data in enumerate(st.session_state.extracted_data):
                                if data['page_number'] == result['page_number']:
                                    existing_index = i
                                    break
                            
                            page_data = {
                                'page_number': result['page_number'],
                                'date': date_value,
                                'invoice_number': invoice_value,
                                'amount': amount_value
                            }
                            
                            if existing_index is not None:
                                st.session_state.extracted_data[existing_index] = page_data
                            else:
                                st.session_state.extracted_data.append(page_data)
                            
                            st.success(f"✅ บันทึกข้อมูลหน้า {result['page_number']} เรียบร้อย!")
            
            # Step 4: Create Excel
            if st.session_state.extracted_data:
                st.header("💾 Step 4: สร้างไฟล์ Excel")
                
                # แสดงสรุปข้อมูลที่จะกรอกลง Excel
                st.subheader("📋 สรุปข้อมูลที่จะกรอกลง Excel:")
                
                summary_data = []
                for data in st.session_state.extracted_data:
                    summary_data.append({
                        'หน้า': data['page_number'],
                        'วันที่': data['date'],
                        'เลขที่ตามบิล': data['invoice_number'],
                        'ยอดก่อน VAT': data['amount']
                    })
                
                df_summary = pd.DataFrame(summary_data)
                st.dataframe(df_summary, use_container_width=True)
                
                # ปุ่มดาวน์โหลด Excel
                col1, col2 = st.columns([3, 1])
                
                with col2:
                    excel_file = fill_excel_with_data(st.session_state.extracted_data)
                    
                    st.download_button(
                        label="⬇️ ดาวน์โหลด Excel",
                        data=excel_file,
                        file_name=f"Invoice_Data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                
                with col1:
                    st.info(f"📊 พร้อมสร้างไฟล์ Excel ด้วยข้อมูล {len(st.session_state.extracted_data)} รายการ")
    
    else:
        st.info("👆 กรุณาอัปโหลดไฟล์ PDF เพื่อเริ่มต้น")
        
        # แสดง Template Excel สำหรับดาวน์โหลด
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
