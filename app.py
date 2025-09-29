import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import os
import io

# =========================================================
# การตั้งค่า Tesseract สำหรับ Cloud
# =========================================================
try:
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
except Exception:
    pass

def optimize_image_for_ocr(image):
    """ปรับปรุงรูปภาพให้เหมาะสำหรับ OCR โดยเฉพาะตัวเลขและภาษาไทย"""
    try:
        # 1. แปลงเป็น grayscale
        if image.mode != 'L':
            image = image.convert('L')
        
        # 2. เพิ่มขนาดภาพ 2 เท่า (ทำให้ OCR แม่นยำขึ้น)
        width, height = image.size
        image = image.resize((width * 2, height * 2), Image.LANCZOS)
        
        # 3. Auto contrast เพื่อให้ตัวอักษรชัดขึ้น
        image = ImageOps.autocontrast(image, cutoff=3)
        
        # 4. เพิ่มความคมชัดมาก
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(3.0)
        
        # 5. เพิ่ม contrast สูง
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.8)
        
        # 6. Apply sharpen filter
        image = image.filter(ImageFilter.SHARPEN)
        
        # 7. ปรับ brightness เล็กน้อย
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.15)
        
        return image
    except Exception as e:
        st.warning(f"Image optimization warning: {e}")
        return image

def extract_invoice_data_precise(ocr_text):
    """ดึงข้อมูลแบบแม่นยำสูง ตาม pattern ที่แน่นอน"""
    
    # ทำความสะอาด text ก่อน
    clean_text = re.sub(r'\s+', ' ', ocr_text.strip())
    
    result = {
        'date': '',
        'invoice_number': '',
        'amount': '',
        'confidence': 0
    }
    
    # === 1. ดึงวันที่: DD/MM/YY Pattern ===
    date_patterns = [
        r'\b(\d{2}/\d{2}/68)\b',           # XX/XX/68 (ปี 68 ตายตัว)
        r'\b(\d{1,2}/\d{1,2}/68)\b',       # X/XX/68 หรือ XX/X/68
        r'(\d{2}/08/68)',                  # XX/08/68 (เดือน 08)
        r'(\d{2}/\d{2}/\d{2})',            # XX/XX/XX ทั่วไป
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, clean_text)
        if matches:
            # เลือกวันที่ที่เป็น format ถูกต้อง
            for date_match in matches:
                if '08/68' in date_match or len(date_match.split('/')) == 3:
                    result['date'] = date_match
                    result['confidence'] += 30
                    break
            if result['date']:
                break
    
    # === 2. ดึงเลขที่: HH6800XXX Pattern ===
    invoice_patterns = [
        r'\b(HH68004\d{2})\b',            # HH68004XX (ตาม pattern ที่แน่นอน)
        r'\b(HH68005\d{2})\b',            # HH68005XX
        r'\b(HH6800\d{3})\b',             # HH6800XXX
        r'(HH\d{7,8})',                   # HH + ตัวเลข 7-8 หลัก
    ]
    
    for pattern in invoice_patterns:
        matches = re.findall(pattern, clean_text)
        if matches:
            # เลือกเลขที่ที่ยาวที่สุดและตรง pattern
            best_invoice = max(matches, key=len) if matches else ''
            if len(best_invoice) >= 9:  # HH + 7 หลักขึ้นไป
                result['invoice_number'] = best_invoice
                result['confidence'] += 30
                break
    
    # === 3. ดึงยอดเงิน: XXXXX.XX Pattern ===
    # ใช้ข้อมูลจริงเป็นไกด์: 4710.28, 16549.53, 17433.64, etc.
    amount_patterns = [
        # Pattern หลัก: ตัวเลข 4-5 หลัก ตามด้วย .XX
        r'\b(\d{4,5}\.\d{2})\b',
        
        # Pattern สำรอง: มี comma คั่น
        r'\b(\d{1,2},\d{3}\.\d{2})\b',
        
        # Pattern เฉพาะ: หาจากบริบท
        r'(?:มูลค่า|Value|รวม|Total|Net)[^0-9]*?(\d{4,5}\.\d{2})',
        
        # Pattern ใกล้ VAT
        r'(\d{4,5}\.\d{2})\s*(?:บาท|VAT|7\.00)',
        
        # Pattern ทั่วไป: ตัวเลขที่มีจุดทศนิยม 2 ตำแหน่ง
        r'(\d{3,6}\.\d{2})',
    ]
    
    found_amounts = []
    
    for i, pattern in enumerate(amount_patterns):
        matches = re.findall(pattern, clean_text, re.IGNORECASE)
        for match in matches:
            # ทำความสะอาดและตรวจสอบ
            clean_amount = match.replace(',', '')
            try:
                amount_value = float(clean_amount)
                # กรองเฉพาะตัวเลขที่อยู่ในช่วงที่เหมาะสม
                if 1000 <= amount_value <= 50000:  # ช่วงยอดเงินที่สมเหตุสมผล
                    found_amounts.append({
                        'amount': clean_amount,
                        'value': amount_value,
                        'priority': i,  # pattern ที่พบ
                        'raw': match
                    })
            except ValueError:
                continue
    
    # เลือกยอดเงินที่ดีที่สุด
    if found_amounts:
        # เรียงตาม priority ของ pattern และความเหมาะสม
        found_amounts.sort(key=lambda x: (x['priority'], -x['value']))
        best_amount = found_amounts[0]
        
        result['amount'] = best_amount['amount']
        result['confidence'] += 40
    
    return result

def process_pdf_ultra_fast(pdf_bytes):
    """ประมวลผล PDF แบบเร็วและแม่นยำ"""
    temp_file = "temp_pdf.pdf"
    
    try:
        # บันทึกไฟล์ชั่วคราว
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        
        # แสดงสถานะ
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("🔄 กำลังแปลง PDF เป็นรูปภาพ...")
        
        # แปลง PDF ด้วยความละเอียดสูง
        pages = convert_from_path(temp_file, dpi=400, fmt='PNG')
        
        results = []
        total_pages = len(pages)
        
        for i, page in enumerate(pages):
            # อัปเดต progress
            progress = (i + 1) / total_pages
            progress_bar.progress(progress)
            status_text.text(f"📖 กำลังประมวลผลหน้าที่ {i+1}/{total_pages}")
            
            # ปรับปรุงรูปภาพ
            optimized_image = optimize_image_for_ocr(page)
            
            # OCR หลายรูปแบบเพื่อความแม่นยำ
            ocr_texts = []
            
            # Config 1: เน้นตัวเลข
            try:
                text1 = pytesseract.image_to_string(
                    optimized_image,
                    lang="tha+eng",
                    config="--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz/.-:ก-๙ "
                )
                ocr_texts.append(text1)
            except:
                pass
            
            # Config 2: เน้นโครงสร้าง
            try:
                text2 = pytesseract.image_to_string(
                    optimized_image,
                    lang="tha+eng",
                    config="--psm 4 --oem 3"
                )
                ocr_texts.append(text2)
            except:
                pass
            
            # รวม text ทั้งหมด
            combined_text = " ".join(ocr_texts)
            
            # ดึงข้อมูล
            extracted_data = extract_invoice_data_precise(combined_text)
            extracted_data['page_number'] = i + 1
            extracted_data['raw_text'] = combined_text[:500]  # เก็บ text บางส่วนไว้ debug
            
            results.append(extracted_data)
        
        # ลบไฟล์ชั่วคราว
        os.remove(temp_file)
        
        # Clear progress
        progress_bar.empty()
        status_text.empty()
        
        return results
        
    except Exception as e:
        if os.path.exists(temp_file):
            os.remove(temp_file)
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
        return []

def create_final_excel(data_list, filename):
    """สร้างไฟล์ Excel สำเร็จรูป"""
    
    # สร้าง DataFrame
    df_data = []
    for i, data in enumerate(data_list, 1):
        df_data.append({
            'ลำดับ': i,
            'วันที่': data['date'],
            'เลขที่ตามบิล': data['invoice_number'], 
            'ยอดก่อน VAT': data['amount']
        })
    
    df = pd.DataFrame(df_data)
    
    # สร้างไฟล์ Excel
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet หลัก
        df.to_excel(writer, sheet_name='Invoice_Data', index=False)
        
        # Sheet สรุป
        summary_data = {
            'รายการ': [
                'จำนวนใบเสร็จทั้งหมด',
                'ใบเสร็จที่มีวันที่',
                'ใบเสร็จที่มีเลขที่',
                'ใบเสร็จที่มียอดเงิน',
                'ยอดรวมทั้งหมด'
            ],
            'จำนวน/ค่า': [
                len(df_data),
                len([d for d in data_list if d['date']]),
                len([d for d in data_list if d['invoice_number']]),
                len([d for d in data_list if d['amount']]),
                sum([float(d['amount']) for d in data_list if d['amount']]) if any(d['amount'] for d in data_list) else 0
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="Ultra-Fast Invoice Extractor",
        page_icon="⚡",
        layout="wide"
    )
    
    # Header
    st.title("⚡ Ultra-Fast PDF Invoice Extractor")
    st.markdown("**หนึ่งปุ่ม - ได้ Excel เลย | สำหรับใบเสร็จภาษาไทย**")
    st.markdown("---")
    
    # คำอธิบายสั้น ๆ
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("📋 **วิธีใช้:** อัปโหลด PDF → กด 'ประมวลผล' → ดาวน์โหลด Excel ✅")
    
    # Main area
    st.header("📁 อัปโหลดไฟล์ PDF")
    
    uploaded_file = st.file_uploader(
        "เลือกไฟล์ PDF ใบเสร็จ",
        type="pdf",
        help="รองรับไฟล์ PDF หลายหน้า"
    )
    
    if uploaded_file is not None:
        
        # แสดงข้อมูลไฟล์
        col1, col2, col3 = st.columns([1, 1, 2])
        
        with col1:
            file_size = len(uploaded_file.getvalue()) / (1024 * 1024)
            st.metric("ขนาดไฟล์", f"{file_size:.1f} MB")
        
        with col2:
            st.metric("สถานะ", "พร้อมประมวลผล ✅")
        
        with col3:
            # ปุ่มประมวลผลหลัก
            if st.button(
                "🚀 ประมวลผลและสร้าง Excel", 
                type="primary", 
                use_container_width=True,
                help="ประมวลผล PDF และสร้างไฟล์ Excel อัตโนมัติ"
            ):
                
                with st.spinner("⚡ กำลังประมวลผลด้วย AI OCR..."):
                    
                    # ประมวลผล PDF
                    pdf_bytes = uploaded_file.getvalue()
                    results = process_pdf_ultra_fast(pdf_bytes)
                
                if results:
                    st.success(f"✅ ประมวลผลเสร็จสิ้น! พบข้อมูล {len(results)} หน้า")
                    
                    # แสดงตัวอย่างข้อมูล
                    st.subheader("📊 ตัวอย่างข้อมูลที่ดึงได้:")
                    
                    preview_data = []
                    for i, result in enumerate(results[:5], 1):  # แสดง 5 รายการแรก
                        preview_data.append({
                            'หน้า': i,
                            'วันที่': result['date'] or '❌',
                            'เลขที่ตามบิล': result['invoice_number'] or '❌',
                            'ยอดก่อน VAT': result['amount'] or '❌',
                            'ความแม่นยำ': f"{result['confidence']}%"
                        })
                    
                    preview_df = pd.DataFrame(preview_data)
                    st.dataframe(preview_df, use_container_width=True)
                    
                    if len(results) > 5:
                        st.info(f"... และอีก {len(results) - 5} รายการ (ดูทั้งหมดในไฟล์ Excel)")
                    
                    # สถิติสรุป
                    st.subheader("📈 สถิติสรุป:")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        total_pages = len(results)
                        st.metric("จำนวนหน้า", total_pages)
                    
                    with col2:
                        valid_dates = len([r for r in results if r['date']])
                        st.metric("วันที่ถูกต้อง", f"{valid_dates}/{total_pages}")
                    
                    with col3:
                        valid_invoices = len([r for r in results if r['invoice_number']])
                        st.metric("เลขที่ถูกต้อง", f"{valid_invoices}/{total_pages}")
                    
                    with col4:
                        valid_amounts = len([r for r in results if r['amount']])
                        st.metric("ยอดเงินถูกต้อง", f"{valid_amounts}/{total_pages}")
                    
                    # คำนวณยอดรวม
                    try:
                        total_amount = sum([float(r['amount']) for r in results if r['amount']])
                        st.metric("💰 ยอดรวมทั้งหมด", f"{total_amount:,.2f} บาท")
                    except:
                        st.metric("💰 ยอดรวมทั้งหมด", "ไม่สามารถคำนวณได้")
                    
                    # สร้างและดาวน์โหลดไฟล์ Excel
                    st.markdown("---")
                    st.subheader("💾 ดาวน์โหลดไฟล์ Excel")
                    
                    excel_file = create_final_excel(results, uploaded_file.name)
                    
                    col1, col2 = st.columns([3, 1])
                    
                    with col2:
                        st.download_button(
                            label="⬇️ ดาวน์โหลด Excel",
                            data=excel_file,
                            file_name=f"Invoice_Data_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    
                    with col1:
                        st.info("📋 ไฟล์ Excel มี 2 Sheet: 'Invoice_Data' (ข้อมูลหลัก) และ 'Summary' (สรุป)")
                    
                    # Debug section (ซ่อนไว้)
                    with st.expander("🔍 ข้อมูล Debug (สำหรับตรวจสอบ)", expanded=False):
                        for result in results:
                            st.text(f"หน้า {result['page_number']}: Confidence {result['confidence']}%")
                            st.text(f"Raw text (100 ตัวอักษรแรก): {result['raw_text'][:100]}...")
                            st.markdown("---")
                
                else:
                    st.error("❌ ไม่สามารถประมวลผลได้ กรุณาลองใหม่อีกครั้ง")
    
    else:
        # แสดงคำแนะนำ
        st.info("👆 กรุณาอัปโหลดไฟล์ PDF เพื่อเริ่มต้น")
        
        # แสดงตัวอย่างข้อมูลที่ต้องการ
        with st.expander("📋 ตัวอย่างข้อมูลที่โปรแกรมจะดึง", expanded=True):
            sample_data = {
                'ลำดับ': [1, 2, 3, 4, 5],
                'วันที่': ['01/08/68', '02/08/68', '03/08/68', '04/08/68', '05/08/68'],
                'เลขที่ตามบิล': ['HH6800470', 'HH6800474', 'HH6800475', 'HH6800476', 'HH6800478'],
                'ยอดก่อน VAT': ['4710.28', '16549.53', '17433.64', '12910.28', '21648.60']
            }
            
            sample_df = pd.DataFrame(sample_data)
            st.dataframe(sample_df, use_container_width=True)
            
            st.markdown("**คุณสมบัติเด่น:**")
            st.markdown("""
            - ⚡ **ความเร็วสูง** - ประมวลผลอัตโนมัติ 100%
            - 🎯 **ความแม่นยำสูง** - AI OCR เฉพาะงานใบเสร็จไทย  
            - 💡 **ใช้งานง่าย** - กดปุ่มเดียวได้ Excel เลย
            - 📊 **ครบครัน** - สรุปสถิติและตรวจสอบข้อมูล
            """)

if __name__ == "__main__":
    main()
