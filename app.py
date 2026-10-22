import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import io
import unicodedata

st.set_page_config(
    page_title="منصة حصين | حماية وتعمية الوثائق القانونية",
    page_icon="⚖️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; direction: rtl; text-align: right; }
    [data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none !important; }
    .main-header {
        background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 100%);
        padding: 25px 15px; border-radius: 12px; border-bottom: 4px solid #c5a059;
        color: #ffffff; text-align: center; margin-bottom: 20px; box-shadow: 0 6px 15px rgba(0,0,0,0.15);
    }
    .main-header h1 { color: #c5a059; font-weight: 800; font-size: 2rem; margin-bottom: 8px; }
    .main-header p { color: #e0e1dd; font-size: 1rem; margin: 0; }
    .legal-box { background-color: #f8f9fa; border-right: 5px solid #c5a059; padding: 15px; border-radius: 8px; margin-bottom: 15px; color: #1b263b; }
    .disclaimer-box { background-color: #fff8e6; border: 1px solid #f0d78c; border-right: 5px solid #d97706; padding: 15px; border-radius: 8px; font-size: 0.9rem; color: #333; margin-bottom: 15px; }
    .stButton>button, .stDownloadButton>button {
        background-color: #0d1b2a !important; color: #c5a059 !important; border: 1px solid #c5a059 !important;
        font-weight: bold !important; border-radius: 8px !important; padding: 12px 20px !important; width: 100%; margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>⚖️ منصة حصين (Haseen)</h1>
    <p>النظام الآلي المتكامل لتعمية الوثائق وصحف الدعاوى مع الحفاظ على التسلسل الزمني وبنية النص</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="legal-box">
    <b>مرحباً بك:</b> تقوم هذه النسخة بنسخ صحيفة الدعوى أو الوثيقة بكامل هيكلها وترويستها وبنودها، مع تعمية الأسرار والبيانات الحساسة فقط، وتحويل التواريخ إلى <b>رموز تسلسلية زمنية (مثل: [تسلسل زمني 1])</b> لضمان عدم اختلال تسلسل وقائع القضية عند دراستها.
</div>
""", unsafe_allow_html=True)

agree_terms = st.checkbox("أقر بقراءة سياسة الاستخدام وإخلاء المسؤولية، وأوافق على مراجعة المستند المعالج.")

with st.expander("⚙️ خيارات التعمية المتقدمة", expanded=False):
    mask_ids = st.checkbox("تعمية أرقام الهويات، السجلات، والصكوك والوكالات", value=True)
    mask_dates = st.checkbox("تحويل التواريخ إلى تسلسل زمني (لحفظ ترابط وقائع القضية)", value=True)
    mask_names = st.checkbox("تعمية أسماء الأطراف والشركات بـ [طرف محجوب]", value=True)
    mask_phones = st.checkbox("تعمية الهواتف والآيبان", value=True)

def smart_redact_document(raw_text, custom_names_input):
    text = unicodedata.normalize('NFKC', raw_text)
    redacted_text = text
    pdf_targets = []
    num = r'[0-9٠-٩۰-۹]'
    
    # 1. الأسماء المخصصة
    if custom_names_input:
        names = [unicodedata.normalize('NFKC', n.strip()) for n in custom_names_input.split('\n') if n.strip()]
        for name in names:
            if name in redacted_text:
                redacted_text = redacted_text.replace(name, "[طرف محجوب]")
            pdf_targets.append(name)

    # 2. تعمية الأسماء الذكية في الصحائف والوكالات
    if mask_names:
        lines = redacted_text.split('\n')
        new_lines = []
        for line in lines:
            if re.search(r'(?:^|\s)(?:مدعي|مدعى عليه|موكّل|وكيل|بن|ابن|بنت)(?:\s|$)', line):
                words = re.findall(r'[\u0600-\u06FF]+', line)
                for w in words:
                    if w not in ['مدعي', 'عليه', 'موكّل', 'وكيل', 'أصالة', 'عن', 'نفسه', 'سعودي', 'سجل', 'مدني']:
                        pdf_targets.append(w)
                new_lines.append("[طرف / صفة محجوبة]")
            else:
                new_lines.append(line)
        redacted_text = '\n'.join(new_lines)

    # 3. التواريخ بأسلوب تسلسلي لحفظ سرد الأحداث
    if mask_dates:
        date_pattern = rf'(?<!{num})({num}{{1,4}}\s*[/-]\s*{num}{{1,2}}\s*[/-]\s*{num}{{1,4}}\s*[مھهـ]?)(?!{num})'
        date_counter = 1
        for match in re.findall(date_pattern, redacted_text):
            full_date = match[0] if isinstance(match, tuple) else match
            pdf_targets.append(full_date)
            redacted_text = redacted_text.replace(full_date, f"[تاريخ سداسي تسلسلي {date_counter}]")
            date_counter += 1

    # 4. الهويات والصكوك والأرقام الطويلة
    if mask_ids:
        id_pattern = rf'(?<!{num}){num}{{8,15}}(?!{num})'
        for m in re.findall(id_pattern, redacted_text):
            pdf_targets.append(m)
        redacted_text = re.sub(id_pattern, '[رقم مرجعي محجوب]', redacted_text)

    # 5. الهواتف والآيبان
    if mask_phones:
        phone_pattern = rf'(?<!{num})(?:05|٠٥|9665|٩٦٦٥){num}{{8}}(?!{num})'
        iban_pattern = rf'SA{num}{{22}}\b'
        for m in re.findall(phone_pattern, redacted_text) + re.findall(iban_pattern, redacted_text):
            pdf_targets.append(m)
        redacted_text = re.sub(phone_pattern, '[هاتف محجوب]', redacted_text)
        redacted_text = re.sub(iban_pattern, '[آيبان محجوب]', redacted_text)

    return redacted_text, pdf_targets

def generate_word_docx(processed_text):
    doc = Document()
    header_para = doc.add_paragraph()
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_para.add_run("صحيفة دعوى / وثيقة معالجة ومحمية (منصة حصين)")
    header_run.bold = True
    header_run.font.size = Pt(13)
    header_run.font.color.rgb = RGBColor(197, 160, 89)
    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for line in processed_text.split('\n'):
        if line.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(line)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def redact_pdf_directly(pdf_bytes, pdf_targets):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    clean_targets = [unicodedata.normalize('NFKC', t.strip()) for t in set(pdf_targets) if len(t.strip()) >= 2]
    
    for page in doc:
        words = page.get_text("words")
        for w in words:
            rect = fitz.Rect(w[:4])
            norm_word = unicodedata.normalize('NFKC', w[4])
            for target in clean_targets:
                if len(target) >= 3 and (target in norm_word or norm_word in target):
                    page.add_redact_annot(rect, fill=(0, 0, 0))
                    break
        page.apply_redactions()
        
    pdf_buffer = io.BytesIO()
    doc.save(pdf_buffer)
    pdf_buffer.seek(0)
    return pdf_buffer

if agree_terms:
    custom_names = st.text_area("أضف أسماء إضافية لطمسها (إن وجدت):", placeholder="شركة كذا\nمحمد فلان", height=70)
    uploaded_file = st.file_uploader("رفع صحيفة الدعوى أو المستند القضائي (PDF):", type=["pdf"])
    
    if uploaded_file is not None:
        pdf_bytes = uploaded_file.read()
        with st.spinner("جاري معالجة المستند والحفاظ على بنية الصحيفة والتسلسل الزمني..."):
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            full_raw_text = ""
            for page in doc:
                full_raw_text += page.get_text("text") + "\n"
                
            processed_text, pdf_targets = smart_redact_document(full_raw_text, custom_names)
            
        st.success("✅ تمت معالجة صحيفة الدعوى بنجاح مع الحفاظ على تسلسل الوقائع وهيكل المستند!")
        
        with st.expander("📄 معاينة النص المهيكل والمعالج", expanded=True):
            final_edited_text = st.text_area("التعديل اليدوي النهائي:", value=processed_text, height=300)
            
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            redacted_pdf_file = redact_pdf_directly(pdf_bytes, pdf_targets)
            st.download_button("📕 تحميل PDF المطموس", data=redacted_pdf_file, file_name=f"Haseen_Redacted_{uploaded_file.name}", mime="application/pdf")
        with col2:
            word_file = generate_word_docx(final_edited_text)
            st.download_button("📄 تحميل Word منسق", data=word_file, file_name=f"Haseen_Protected.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.warning("يرجى الموافقة على إخلاء المسؤولية أعلاه لتفعيل الرفع.")
