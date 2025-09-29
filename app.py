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
    """ปรับปรุงรูปภาพให้เหมาะสำหรับ OCR"""
    try:
        # 1. แปลงเป็น grayscale
        if image.mode != 'L':
            image = image.convert('L')
        
        # 2. เพิ่มขนาดภาพ 2 เท่า
        width, height = image.size
        image = image.resize((width * 2, height * 2), Image.LANCZOS)
        
        # 3. Auto contrast
        image = ImageOps.autocontrast(image, cutoff=3)
        
        # 4. เพิ่มความคมชัด
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(4.0)
        
        # 5. เพิ่ม contrast
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(3.5)
        
        # 6. Apply UnsharpMask filter
        image = image.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
        
        # 7. ปรับ brightness
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.3)
        
        # 8. Add binary threshold
        image = image.point(lambda x: 0 if x < 130 else 255)
        
        return image
    except Exception as e:
        st.warning(f"Image optimization warning: {e}")
        return image

def optimize_image_for_display(image):
    """ปรับปรุงรูปภาพให้ชัดเจนสำหรับการแสดงผล"""
    try:
        max_width = 800
        ratio = max_width / image.width
        new_height = int(image.height * ratio)
        image = image.resize((max_width, new_height), Image.LANCZOS)
        
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)
        
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(1.5)
        
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.2)
        
        return image
    except Exception as e:
        st.warning(f"Image display optimization warning: {e}")
        return image

def extract_invoice_data_precise(ocr_text):
    """ดึงข้อมูลแบบแม่นยำ"""
    lines = [line.strip() for line in ocr_text.split('\n') if line.strip()]
    clean_text = ' '.join(lines)
    
    result = {
        'date': '',
        'invoice_number': '',
        'amount': '',
        'confidence': 0,
        'debug_matches': {}
    }
    
    date_patterns = [r'(\d{2}/\d{2}/\d{2})', r'(\d{1,2}/\d{1,2}/\d{4})']  # Generalized date patterns
    for line in lines:
        if 'Date' in line or 'วันที่' in line or 'Date' in line.upper():
            for pattern in date_patterns:
                match = re.search(pattern, line)
                if match:
                    result['date'] = match.group(1)
                    result['confidence'] += 30
                    result['debug_matches']['date_line'] = line
                    break
            if result['date']:
                break
    if not result['date']:
        for pattern in date_patterns:
            matches = re.findall(pattern, clean_text)
            if matches:
                result['date'] = matches[0]
                result['confidence'] += 20
                break
    
    invoice_patterns = [r'(INV-\d{6})', r'(HH\d{7})', r'(\w{2}\d{6})']  # Generalized invoice patterns
    for line in lines:
        if 'No.' in line or 'เลขที่' in line or 'No.' in line.upper():
            for pattern in invoice_patterns:
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    result['invoice_number'] = match.group(1)
                    result['confidence'] += 30
                    result['debug_matches']['invoice_line'] = line
                    break
            if result['invoice_number']:
                break
    if not result['invoice_number']:
        for pattern in invoice_patterns:
            matches = re.findall(pattern, clean_text, re.IGNORECASE)
            if matches:
                result['invoice_number'] = matches[0]
                result['confidence'] += 20
                break
    
    amount_patterns = [
        r'Amount\s*([,\d]+\.\d{2})',
        r'มูลค่า\s*([,\d]+\.\d{2})',
        r'([,\d]+\.\d{2})\s*(?:USD|THB|JPY|CNY)?',
    ]
    found_amount = False
    for line in lines:
        for pattern in amount_patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                raw_amount = match.group(1)
                clean_amount = raw_amount.replace(',', '')
                try:
                    amount_value = float(clean_amount)
                    if 100 <= amount_value <= 100000:  # Adjusted range
                        result['amount'] = clean_amount
                        result['confidence'] += 40
                        result['debug_matches']['amount_line'] = line
                        found_amount = True
                        break
                except ValueError:
                    continue
        if found_amount:
            break
    if not found_amount:
        fallback_patterns = [r'\b(\d{3,6}\.\d{2})\b']
        matches = re.findall(fallback_patterns[0], clean_text)
        if matches:
            clean_amount = matches[0].replace(',', '')
            try:
                amount_value = float(clean_amount)
                if 100 <= amount_value <= 100000:
                    result['amount'] = clean_amount
                    result['confidence'] += 20
            except ValueError:
                pass
    
    result['debug_matches']['lines_sample'] = lines[:10]
    return result

def process_pdf_ultra_fast(pdf_bytes, selected_langs):
    """ประมวลผล PDF ด้วยภาษาที่เลือก"""
    temp_file = "temp_pdf.pdf"
    
    try:
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("🔄 กำลังแปลง PDF เป็นรูปภาพ...")
        
        pages = convert_from_path(temp_file, dpi=400, fmt='PNG')
        
        results = []
        total_pages = len(pages)
        
        for i, page in enumerate(pages):
            progress = (i + 1) / total_pages
            progress_bar.progress(progress)
            status_text.text(f"📖 กำลังประมวลผลหน้าที่ {i+1}/{total_pages}")
            
            optimized_image = optimize_image_for_ocr(page)
            ocr_texts = []
            
            # Config 1: เน้นตัวเลขและภาษาที่เลือก
            try:
                text1 = pytesseract.image_to_string(
                    optimized_image,
                    lang="+".join(selected_langs),
                    config="--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz/.-:ก-๙ "
                )
                ocr_texts.append(text1)
            except:
                pass
            
            # Config 2: เน้นโครงสร้าง
            try:
                text2 = pytesseract.image_to_string(
                    optimized_image,
                    lang="+".join(selected_langs),
                    config="--psm 4 --oem 3"
                )
                ocr_texts.append(text2)
            except:
                pass
            
            combined_text = "\n".join(ocr_texts)
            extracted_data = extract_invoice_data_precise(combined_text)
            extracted_data['page_number'] = i + 1
            extracted_data['raw_text'] = combined_text
            
            results.append(extracted_data)
        
        os.remove(temp_file)
        progress_bar.empty()
        status_text.empty()
        
        return results, pages
        
    except Exception as e:
        if os.path.exists(temp_file):
            os.remove(temp_file)
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
        return [], []

def create_final_excel(data_list, filename):
    """สร้างไฟล์ Excel"""
    df_data = []
    for i, data in enumerate(data_list, 1):
        df_data.append({
            'ลำดับ': i,
            'วันที่': data['date'],
            'เลขที่ตามบิล': data['invoice_number'], 
            'ยอดก่อน VAT': data['amount']
        })
    
    df = pd.DataFrame(df_data)
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Invoice_Data', index=False)
        summary_data = {
            'รายการ': ['จำนวนใบเสร็จทั้งหมด', 'ใบเสร็จที่มีวันที่', 'ใบเสร็จที่มีเลขที่', 'ใบเสร็จที่มียอดเงิน', 'ยอดรวมทั้งหมด'],
            'จำนวน/ค่า': [
                len(df_data),
                len([d for d in data_list if d['date']]),
                len([d for d in data_list if d['invoice_number']]),
                len([d for d in data_list if d['amount']]),
                sum([float(d['amount']) if d['amount'] else 0 for d in data_list])
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Ultra-Fast Invoice Extractor", page_icon="⚡", layout="wide")
    st.title("⚡ Ultra-Fast PDF Invoice Extractor")
    st.markdown("**หนึ่งปุ่ม - ได้ Excel เลย | รองรับหลายภาษา**")
    st.markdown("---")
    
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.info("📋 **วิธีใช้:** เลือกภาษา → อัปโหลด PDF → กด 'ประมวลผล' → ตรวจสอบข้อมูลและภาพ → ดาวน์โหลด Excel ✅")
    
    st.header("🌐 เลือกภาษาที่ต้องการ")
    language_options = {
        "Thai": "tha",
        "English": "eng",
        "Chinese (Simplified)": "chi_sim",
        "Japanese": "jpn",
        "French": "fra",
        "Spanish": "spa"
    }
    selected_langs = st.multiselect(
        "เลือกภาษา (เลือกได้หลายภาษา)", 
        options=list(language_options.keys()), 
        default=["Thai", "English"],
        help="เลือกภาษาที่ใช้ในเอกสารของคุณ"
    )
    selected_lang_codes = [language_options[lang] for lang in selected_langs]
    
    st.header("📁 อัปโหลดไฟล์ PDF")
    uploaded_file = st.file_uploader("เลือกไฟล์ PDF ใบเสร็จ", type="pdf")
    
    if uploaded_file is not None:
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            file_size = len(uploaded_file.getvalue()) / (1024 * 1024)
            st.metric("ขนาดไฟล์", f"{file_size:.1f} MB")
        with col2:
            st.metric("สถานะ", "พร้อมประมวลผล ✅")
        with col3:
            if st.button("🚀 ประมวลผลและสร้าง Excel", type="primary", use_container_width=True):
                with st.spinner("⚡ กำลังประมวลผลด้วย AI OCR..."):
                    pdf_bytes = uploaded_file.getvalue()
                    results, page_images = process_pdf_ultra_fast(pdf_bytes, selected_lang_codes)
                
                if results:
                    st.success(f"✅ ประมวลผลเสร็จสิ้น! พบข้อมูล {len(results)} หน้า")
                    st.subheader("📊 ตัวอย่างข้อมูลที่ดึงได้:")
                    preview_data = []
                    for i, result in enumerate(results[:5], 1):
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
                    
                    total_amount = sum([float(r['amount']) if r['amount'] else 0 for r in results])
                    st.metric("💰 ยอดรวมทั้งหมด", f"{total_amount:,.2f} บาท")
                    
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
                    
                    st.subheader("🔍 การตรวจสอบข้อมูลที่ดึงได้พร้อมภาพเอกสาร")
                    for idx, result in enumerate(results):
                        with st.expander(f"หน้า {idx+1} - ผลลัพธ์: {result['date']} | {result['invoice_number']} | {result['amount']} | ความแม่นยำ {result['confidence']}%"):
                            optimized_display_image = optimize_image_for_display(page_images[idx])
                            st.image(optimized_display_image, caption=f"ภาพเอกสารหน้า {idx+1} (ปรับปรุงความชัด)", use_column_width=True)
                            st.text_area(f"Raw OCR Text หน้า {idx+1}:", result.get('raw_text', ''), height=300)
                            if 'debug_matches' in result:
                                st.write("Debug Matches:")
                                st.json(result['debug_matches'])
                    
                    st.warning("⚠️ กรุณาตรวจสอบข้อมูลและภาพเอกสารข้างต้น หากไม่ถูกต้องให้แจ้งเพื่อปรับปรุง pattern")
                
                else:
                    st.error("❌ ไม่สามารถประมวลผลได้ กรุณาลองใหม่อีกครั้ง")
    
    else:
        st.info("👆 กรุณาอัปโหลดไฟล์ PDF เพื่อเริ่มต้น")
        with st.expander("📋 ตัวอย่างข้อมูลที่โปรแกรมจะดึง", expanded=True):
            sample_data = {
                'ลำดับ': [1, 2, 3],
                'วันที่': ['01/09/25', '02/09/25', '03/09/25'],
                'เลขที่ตามบิล': ['INV-123456', 'INV-123457', 'INV-123458'],
                'ยอดก่อน VAT': ['1500.50', '2500.75', '3000.00']
            }
            sample_df = pd.DataFrame(sample_data)
            st.dataframe(sample_df, use_container_width=True)
            st.markdown("**คุณสมบัติเด่น:**")
            st.markdown("""
            - ⚡ **ความเร็วสูง** - ประมวลผลอัตโนมัติ 100%
            - 🎯 **รองรับหลายภาษา** - OCR สำหรับภาษาที่เลือก
            - 💡 **ใช้งานง่าย** - กดปุ่มเดียวได้ Excel เลย
            - 📊 **ครบครัน** - สรุปสถิติและตรวจสอบข้อมูลพร้อมภาพเอกสารชัดเจน
            """)

if __name__ == "__main__":
    main()
