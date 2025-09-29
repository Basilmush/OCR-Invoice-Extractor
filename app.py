import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageEnhance
import os
import io

# =========================================================
# การตั้งค่า Tesseract Path สำหรับ Cloud Server
# =========================================================
try:
    # กำหนด Tesseract Path สำหรับ Linux/Cloud Server
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
except Exception:
    pass


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
        
        # แปลง PDF เป็นรูปภาพ (Poppler จะถูกหา Path ได้อัตโนมัติจาก packages.txt)
        pages = convert_from_path(temp_file, dpi=400)
        
        ocr_results = []
        
        for i, page in enumerate(pages):
            st.info(f"📖 กำลังทำ OCR หน้าที่ {i+1}/{len(pages)}...")
            
            # ปรับปรุงคุณภาพรูปภาพ
            enhanced_page = enhance_image_for_ocr(page)
            
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
    
    # --- 1. การดึงเลขที่ HH (แข็งแกร่งกว่าเดิม) ---
    # ค้นหา HHxxxxxx ในบรรทัดแรกๆ ของเอกสาร
    invoice_pattern = r'เลขที\s*(?:No\.)?\s*([H]\w{6,8})'
    invoice_matches = re.search(invoice_pattern, text, re.IGNORECASE)
    if not invoice_matches:
        # Fallback 1: หา HHxxxxxx ที่ไหนก็ได้
        invoice_pattern = r'(HH\d{6,8})'
        invoice_matches = re.search(invoice_pattern, text)

    if invoice_matches:
        data['invoice_number'] = invoice_matches.group(1)
        data['raw_matches']['invoices_found'] = [data['invoice_number']]
    
    # --- 2. การดึงวันที่ (แข็งแกร่งกว่าเดิม) ---
    # ค้นหา วันที่ หรือ Date แล้วตามด้วย DD/MM/YY
    date_pattern = r'(?:วันที|Date)\s*(\d{1,2}/\d{1,2}/\d{2,4})'
    date_matches = re.search(date_pattern, text, re.IGNORECASE)
    if date_matches:
        data['date'] = date_matches.group(1)
        data['raw_matches']['dates_found'] = [data['date']]
    
    # --- 3. การดึงยอดก่อน VAT (มูลค่าสินค้า) - เน้นตำแหน่งที่แน่นอน ---
    
    # Regex ขั้นสูงสุด: ใช้การค้นหาแบบยืดหยุ่นสำหรับคำนำหน้าและดึงตัวเลขที่ตามมา
    # ค้นหาคำที่ขึ้นต้นด้วย มูลค่าสินค้า หรือ Product Value โดยยอมให้มีตัวอักษรผิดเพี้ยน [มม]*
    amount_pattern_fuzzy = r"(?:[มม]*ูลค่าสินค้า|Product\s*Value)\s*[.,:\s\n\r]*\s*([,\d]+\.\d{2})"
    
    # Regex สำรอง (Fallback): ค้นหาตัวเลขที่อยู่หลัง 'หักส่วนลด' และก่อน 'จำนวนภาษีมูลค่าเพิ่ม'
    # ใช้ re.DOTALL เพื่อให้ '.' รวมการขึ้นบรรทัดใหม่ด้วย
    amount_pattern_deep_fallback = r"(?:หักส่วนลด|Less Discount)(?:.|\n)*?([,\d]+\.\d{2})\s*(?:จำนวนภาษีมูลค่าเพิ่ม|7.00 %)"

    
    amount_match = re.search(amount_pattern_fuzzy, text, re.IGNORECASE)
    
    # ถ้าหาแบบ Fuzzy ไม่เจอ ให้ลองหาแบบ Deep Fallback
    if not amount_match:
        amount_match = re.search(amount_pattern_deep_fallback, text, re.IGNORECASE | re.DOTALL)

    if amount_match:
        # group(1) คือตัวเลขที่ถูกจับกลุ่ม
        raw_amount = amount_match.group(1).replace(',', '')
        try:
            # แปลงเป็น float เพื่อให้แน่ใจว่าเป็นตัวเลข
            data['amount'] = f"{float(raw_amount):.2f}"
        except ValueError:
            data['amount'] = raw_amount
            
    data['raw_matches']['amounts_found'] = [data['amount']]
    
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
    # สร้าง DataFrame จาก List of Dictionaries
    df_data = pd.DataFrame(data_list)
    
    # จัดเรียงตามเลขหน้า
    df_data = df_data.sort_values(by='page_number').reset_index(drop=True)
    
    # แปลงเป็น Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # เลือกเฉพาะคอลัมน์ที่ต้องการ (วันที่, เลขที่ตามบิล, ยอดก่อน VAT)
        df_to_excel = df_data[['date', 'invoice_number', 'amount']].copy()
        df_to_excel.insert(0, 'ลำดับ', df_to_excel.index + 1)
        
        # เปลี่ยนชื่อคอลัมน์ให้ตรงกับที่ต้องการ
        df_to_excel.columns = ['ลำดับ', 'วันที่', 'เลขที่ตามบิล', 'ยอดก่อน VAT']
        
        df_to_excel.to_excel(writer, index=False, sheet_name='Invoice_Data')
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="PDF OCR Extractor",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("🔍 PDF OCR Checker & Excel Data Filler")
    st.markdown("**(สำหรับใบเสร็จ บริษัท ธนารัตน์ปิยะปิโตรเลียม จำกัด)**")
    st.markdown("---")
    
    # Initialize session state
    if 'ocr_results' not in st.session_state:
        st.session_state.ocr_results = []
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = []
    
    # Sidebar
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
        - **วันที่** (กรอกในช่อง 'วันที่')
        - **เลขที่ตามบิล** (กรอกในช่อง 'เลขที่ตามบิล')
        - **มูลค่าสินค้า** (กรอกในช่อง 'ยอดก่อน VAT')
        """)
    
    # --- Step 1 & 2: Upload and Convert ---
    st.header("1. 📁 อัปโหลดไฟล์ PDF")
    uploaded_file = st.file_uploader(
        "เลือกไฟล์ PDF",
        type="pdf",
        help="อัปโหลดไฟล์ใบเสร็จ PDF ที่รวมหลายหน้าได้"
    )
    
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
            
            # --- Step 3: Display and Edit Results ---
            st.header("2. 👁️ ตรวจสอบและแก้ไขข้อมูล (จำเป็น)")
            st.warning("⚠️ โปรดตรวจสอบข้อมูลที่ดึงมาโดย OCR ในช่องสีเหลือง แล้วกด 'บันทึก' ทุกหน้า")
            
            # Loop เพื่อแสดงผลลัพธ์แต่ละหน้า
            for result in st.session_state.ocr_results:
                page_key = result['page_number']
                
                # ดึงข้อมูลจาก OCR (ทุกครั้งที่หน้าถูกสร้าง)
                extracted = extract_data_from_ocr_text(result['ocr_text'])
                
                # ค้นหาค่าที่บันทึกไว้แล้ว (ถ้ามี)
                saved_data = next((d for d in st.session_state.extracted_data if d['page_number'] == page_key), None)
                
                with st.expander(f"📄 ใบเสร็จหน้าที่ {page_key}", expanded=False):
                    
                    st.markdown(f"**สถานะ:** {'💾 บันทึกแล้ว' if saved_data else '✏️ รอการบันทึก/แก้ไข'}")
                    
                    col1, col2 = st.columns([1, 1])
                    
                    with col1:
                        st.subheader("📝 ข้อมูลที่ดึงได้:")
                        
                        # ใช้ค่าที่บันทึกไว้ (ถ้ามี) หรือใช้ค่าจาก OCR ใหม่
                        initial_date = saved_data['date'] if saved_data else extracted['date']
                        initial_invoice = saved_data['invoice_number'] if saved_data else extracted['invoice_number']
                        initial_amount = saved_data['amount'] if saved_data else extracted['amount']
                        
                        # --- Input Fields for User Correction ---
                        date_value = st.text_input(
                            "📅 วันที่:",
                            value=initial_date,
                            key=f"date_{page_key}"
                        )
                        
                        invoice_value = st.text_input(
                            "🔢 เลขที่ตามบิล:",
                            value=initial_invoice,
                            key=f"invoice_{page_key}"
                        )
                        
                        amount_value = st.text_input(
                            "💰 ยอดก่อน VAT:",
                            value=initial_amount,
                            key=f"amount_{page_key}"
                        )
                        
                        # ปุ่มบันทึกข้อมูล
                        if st.button(f"💾 บันทึก/อัปเดตข้อมูลหน้า {page_key}", key=f"save_{page_key}", type="primary", use_container_width=True):
                            
                            # ตรวจสอบว่ามีข้อมูลหน้านี้อยู่แล้วหรือไม่
                            existing_index = next((i for i, data in enumerate(st.session_state.extracted_data) if data['page_number'] == page_key), None)
                            
                            page_data = {
                                'page_number': page_key,
                                'date': date_value,
                                'invoice_number': invoice_value,
                                'amount': amount_value
                            }
                            
                            if existing_index is not None:
                                st.session_state.extracted_data[existing_index] = page_data
                            else:
                                st.session_state.extracted_data.append(page_data)
                            
                            # บังคับให้ Streamlit รันซ้ำเพื่ออัปเดตสถานะ
                            st.rerun() 
                    
                    with col2:
                        st.subheader("🖼️ รูปภาพ OCR & Text (ต้นฉบับ):")
                        st.image(result['image'], use_container_width=True)
                        
                        with st.expander("📝 ข้อความ OCR ดิบ (Raw Text)"):
                             st.text_area(
                                "OCR Text:",
                                result['ocr_text'],
                                height=250,
                                key=f"ocr_text_{page_key}"
                            )
            
            # --- Step 4: Create Excel ---
            if st.session_state.extracted_data:
                st.header("3. 💾 ดาวน์โหลดไฟล์ Excel")
                
                # แสดงสรุปข้อมูลที่จะกรอกลง Excel
                df_summary = pd.DataFrame(st.session_state.extracted_data)
                df_summary = df_summary.sort_values(by='page_number').reset_index(drop=True)
                
                st.subheader("📋 สรุปข้อมูลที่ถูกบันทึกแล้ว:")
                st.dataframe(df_summary, use_container_width=True, height=300)
                
                st.info(f"📊 พร้อมสร้างไฟล์ Excel ด้วยข้อมูล {len(st.session_state.extracted_data)} รายการ")
                
                # ปุ่มดาวน์โหลด Excel
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
