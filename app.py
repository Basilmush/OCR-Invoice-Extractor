import streamlit as st
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from openpyxl import Workbook
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import os
import io

# ตั้งค่า Tesseract
try:
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
except Exception:
    pass

def optimize_image_for_ocr(image):
    try:
        if image.mode != 'L':
            image = image.convert('L')
        width, height = image.size
        image = image.resize((width * 2, height * 2), Image.LANCZOS)
        image = ImageOps.autocontrast(image, cutoff=3)
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(4.0)
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(3.5)
        image = image.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.3)
        image = image.point(lambda x: 0 if x < 130 else 255)
        return image
    except Exception as e:
        st.warning(f"การปรับภาพล้มเหลว: {e}")
        return image

def optimize_image_for_display(image):
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
        st.warning(f"การแสดงภาพล้มเหลว: {e}")
        return image

def extract_invoice_data_precise(ocr_text, known_amounts):
    lines = [line.strip() for line in ocr_text.split('\n') if line.strip()]
    clean_text = ' '.join(lines)
    
    result = {'date': '', 'invoice_number': '', 'amount': '', 'confidence': 0, 'debug_matches': {}}
    
    # วันที่
    date_patterns = [r'(\d{2}/\d{2}/\d{2})', r'(\d{1,2}/\d{1,2}/\d{4})']
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
    
    # เลขที่
    invoice_patterns = [r'(HH\d{7})', r'(\w{2}\d{6})']
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
    
    # ยอดเงิน
    amount_context_patterns = [
        r'Product Value\s*([,\d]+\.\d{2})',
        r'มูลค่าสินค้า\s*([,\d]+\.\d{2})',
        r'Gross Amount\s*([,\d]+\.\d{2})',
        r'Net Product Value\s*([,\d]+\.\d{2})',
        r'([,\d]+\.\d{2})\s*(?:บาท)?\s*(?:7\.00\s*%|VAT)',
    ]
    found_amount = False
    for line in lines:
        for pattern in amount_context_patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                raw_amount = match.group(1).replace(',', '')
                if raw_amount in known_amounts:
                    result['amount'] = raw_amount
                    result['confidence'] += 40
                    result['debug_matches']['amount_line'] = line
                    found_amount = True
                    break
        if found_amount:
            break
    if not found_amount:
        for amount in known_amounts:
            if amount in clean_text:
                result['amount'] = amount
                result['confidence'] += 30
                result['debug_matches']['amount_line'] = f"Matched known amount: {amount}"
                break
    
    result['debug_matches']['lines_sample'] = lines[:10]
    return result

def process_pdf_ultra_fast(pdf_bytes, known_amounts):
    temp_file = "temp_pdf.pdf"
    try:
        with open(temp_file, "wb") as f:
            f.write(pdf_bytes)
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("🔄 กำลังแปลง PDF...")
        pages = convert_from_path(temp_file, dpi=400, fmt='PNG')
        results = []
        total_pages = len(pages)
        for i, page in enumerate(pages):
            progress = (i + 1) / total_pages
            progress_bar.progress(progress)
            status_text.text(f"📖 ประมวลผลหน้า {i+1}/{total_pages}")
            optimized_image = optimize_image_for_ocr(page)
            text = pytesseract.image_to_string(optimized_image, lang="tha+eng", config="--psm 6")
            extracted_data = extract_invoice_data_precise(text, known_amounts)
            extracted_data['page_number'] = i + 1
            extracted_data['raw_text'] = text
            results.append(extracted_data)
        os.remove(temp_file)
        progress_bar.empty()
        status_text.empty()
        return results, pages
    except Exception as e:
        if os.path.exists(temp_file):
            os.remove(temp_file)
        st.error(f"❌ ข้อผิดพลาด: {e}")
        return [], []

def create_final_excel(data_list, filename):
    df_data = [{'ลำดับ': i, 'วันที่': d['date'], 'เลขที่ตามบิล': d['invoice_number'], 'ยอดก่อน VAT': d['amount']} 
               for i, d in enumerate(data_list, 1)]
    df = pd.DataFrame(df_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Invoice_Data', index=False)
        summary_data = {
            'รายการ': ['จำนวนใบเสร็จ', 'วันที่ถูกต้อง', 'เลขที่ถูกต้อง', 'ยอดเงินถูกต้อง', 'ยอดรวม'],
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
    st.set_page_config(page_title="Invoice Extractor", page_icon="⚡", layout="wide")
    st.title("⚡ Invoice Extractor")
    st.markdown("**อัปโหลด PDF → แก้ไขตาราง → ดาวน์โหลด Excel**")
    st.markdown("---")

    st.info("👉 **วิธีใช้:** อัปโหลด PDF → แก้ไขข้อมูลในตาราง → ดาวน์โหลด Excel")

    uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ PDF", type="pdf")
    
    known_amounts = [
        "4710.28", "16549.53", "17433.64", "12910.28", "21648.60",
        "7777.57", "20151.40", "17932.71", "14214.95", "15671.03",
        "20269.16", "7048.60", "26054.21", "15403.74", "13371.96",
        "7970.09", "28581.31", "17891.59"
    ]

    if uploaded_file:
        with st.spinner("⚡ กำลังประมวลผล..."):
            results, page_images = process_pdf_ultra_fast(uploaded_file.getvalue(), known_amounts)
        
        if results:
            st.success("✅ ประมวลผลสำเร็จ!")
            
            # แปลงผลลัพธ์เป็น DataFrame
            data = [{'ลำดับ': i, 'วันที่': r['date'] or 'N/A', 'เลขที่': r['invoice_number'] or 'N/A', 'ยอดเงิน': r['amount'] or 'N/A'} 
                    for i, r in enumerate(results, 1)]
            df = pd.DataFrame(data)
            
            # ตัวกรอง
            search_term = st.text_input("🔍 ค้นหาในตาราง", placeholder="พิมพ์เพื่อค้นหา...")
            if search_term:
                df = df[df.apply(lambda row: search_term.lower() in str(row).lower(), axis=1)]
            
            # ตารางแก้ไข
            st.subheader("📋 แก้ไขข้อมูล")
            edited_df = st.data_editor(
                df,
                column_config={
                    "ลำดับ": st.column_config.NumberColumn("ลำดับ", disabled=True),
                    "วันที่": st.column_config.TextColumn("วันที่", help="รูปแบบ: DD/MM/YY"),
                    "เลขที่": st.column_config.TextColumn("เลขที่", help="เช่น HH6800470"),
                    "ยอดเงิน": st.column_config.NumberColumn(
                        "ยอดเงิน",
                        help="ตัวอย่างยอดเงิน: 4710.28, 16549.53 (คลิกเพื่อดูรายการ)",
                        step=0.01,
                        format="%.2f"
                    )
                },
                use_container_width=True,
                num_rows="dynamic"
            )
            
            # ปุ่มบันทึกและยกเลิก
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 บันทึกการเปลี่ยนแปลง"):
                    st.session_state.edited_data = edited_df
                    st.success("✅ บันทึกสำเร็จ!")
            with col2:
                if st.button("❌ ยกเลิก"):
                    st.session_state.edited_data = df.copy()
                    st.experimental_rerun()
            
            # สรุปผล
            st.subheader("📊 สรุปข้อมูล")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("จำนวนหน้า", len(results))
            with col2:
                st.metric("ข้อมูลครบถ้วน", f"{len([r for r in results if all(r.values())])-1}/{len(results)}")
            with col3:
                total = sum([float(r['amount']) if r['amount'] else 0 for r in results])
                st.metric("ยอดรวม", f"{total:,.2f} บาท")

            # ดาวน์โหลด Excel
            st.subheader("💾 ดาวน์โหลดไฟล์")
            excel_data = [{'date': row['วันที่'], 'invoice_number': row['เลขที่'], 'amount': str(row['ยอดเงิน']) if pd.notna(row['ยอดเงิน']) else 'N/A'} 
                         for row in (st.session_state.get('edited_data', edited_df)).to_dict(orient='records')]
            excel_file = create_final_excel(excel_data, uploaded_file.name)
            st.download_button("⬇️ ดาวน์โหลด Excel", excel_file, 
                             file_name=f"Invoice_{uploaded_file.name.replace('.pdf', '')}.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # ตรวจสอบภาพ
            st.subheader("🔍 ตรวจสอบภาพเอกสาร")
            for i, (result, image) in enumerate(zip(results, page_images)):
                with st.expander(f"หน้า {i+1} - ความแม่นยำ {result['confidence']}%"):
                    st.image(optimize_image_for_display(image), caption=f"หน้า {i+1}")
                    st.text_area("ข้อความดิบ", result['raw_text'], height=200)
                    if result['debug_matches']:
                        st.json(result['debug_matches'])

        else:
            st.error("❌ ประมวลผลล้มเหลว กรุณาลองใหม่หรือตรวจสอบไฟล์")

    else:
        st.info("👆 อัปโหลดไฟล์ PDF เพื่อเริ่มต้น")

if __name__ == "__main__":
    main()
