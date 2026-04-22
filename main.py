import streamlit as st  
import tempfile  
import os  
import subprocess  
from pathlib import Path 
from pdf2docx import Converter  
import shutil 
from PIL import Image
from PyPDF2 import PdfMerger   

# ---------------------------------------------------
# LANGUAGE SYSTEM
# ---------------------------------------------------
LANG = {             
    "en": {                                                                 
        "title": "Convert, merge and transform PDFs, Word files and images in seconds.",
        "subtitle": "No installation required, everything works directly in the browser.",
        "upload": "Upload file 📄",
        "docx_info": "Word → PDF conversion in progress...",
        "pdf_info": "PDF → Word conversion in progress...",
        "done": "Conversion completed!",
        "download_pdf": "⬇️ Download PDF",
        "download_docx": "⬇️ Download Word",
        "error": "Error during conversion.",
        "footer1": "Powered by Streamlit + LibreOffice + Python",
        "footer2": "Created by Alberto Floris",
        "title_main": "PDF ↔ Word",
        "button_lang": "🌐 Italiano",

        # IMAGE → PDF
        "img_title": "Convert images to PDF",
        "img_upload": "Upload one or more images 🖼️",
        "img_info": "Image → PDF conversion in progress...",
        "download_img_pdf": "⬇️ Download PDF",

        # MERGE PDF
        "merge_title": "Merge multiple PDFs",
        "merge_upload": "Upload two or more PDF files 📚",
        "merge_info": "Merging PDF files...",
        "download_merged_pdf": "⬇️ Download merged PDF",

        # DONATION
        "donation_title": "Henkanix grows thanks to small gestures like yours",
        "donation_subtitle": "If you find it useful, you can support the project",
        "donation_button": "💙 Support Henkanix"
    },
    "it": {
        "title": "Converti, unisci e trasforma PDF, Word e immagini in pochi secondi.",
        "subtitle": "Nessuna installazione richiesta, tutto funziona direttamente dal browser.",
        "upload": "Carica file 📄",
        "docx_info": "Conversione Word → PDF in corso...",
        "pdf_info": "Conversione PDF → Word in corso...",
        "done": "Conversione completata!",
        "download_pdf": "⬇️ Scarica PDF",
        "download_docx": "⬇️ Scarica Word",
        "error": "Errore durante la conversione.",
        "footer1": "Powered by Streamlit + LibreOffice + Python",
        "footer2": "Creato da Alberto Floris",
        "title_main": "PDF ↔ Word",
        "button_lang": "🌐 English",

        # IMAGE → PDF
        "img_title": "Converti immagini in PDF",
        "img_upload": "Carica una o più immagini 🖼️",
        "img_info": "Conversione Immagini → PDF in corso...",
        "download_img_pdf": "⬇️ Scarica PDF",

        # MERGE PDF
        "merge_title": "Unisci più PDF",
        "merge_upload": "Carica due o più PDF 📚",
        "merge_info": "Unione dei PDF in corso...",
        "download_merged_pdf": "⬇️ Scarica PDF unito",

        # DONATION
        "donation_title": "Henkanix cresce anche grazie a piccoli gesti come il tuo",
        "donation_subtitle": "Se ti è utile, puoi supportare lo sviluppo del progetto",
        "donation_button": "💙 Supporta Henkanix"
    }
}

# ---------------------------------------------------
# SESSION LANGUAGE
# ---------------------------------------------------
if "lang" not in st.session_state: 
    st.session_state.lang = "en" 

def t(key):   
    return LANG[st.session_state.lang][key]  
                                                 
# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------
st.set_page_config(
    page_title="Henkanix",
    page_icon="✨",
    layout="centered"
)

# ---------------------------------------------------
# LANGUAGE BUTTON
# ---------------------------------------------------
with st.sidebar:   
    current_lang = st.session_state.lang   
    button_label = t("button_lang")  
                                    
    if st.button(button_label):  
        st.session_state.lang = "it" if current_lang == "en" else "en" 
        st.rerun()  

# ---------------------------------------------------
# TITLE
# ---------------------------------------------------
st.markdown(
    """
    <h1 style='text-align:center; font-size:48px; margin-bottom:0;'>
        ✨ Henkanix ✨
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown(
    f"""
    <div style='text-align:center; font-size:20px;
                margin-top:16px; line-height:1.8;
                max-width:780px; margin-left:auto; margin-right:auto;
                font-weight:500;'>
        {t("title")}<br>
        {t("subtitle")}
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("---")

# ---------------------------------------------------
# CSS
# ---------------------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 2rem;}
.stDownloadButton button, .stButton button{
    width:100%;
    border-radius:10px;
    height:48px;
}
.small {
    font-size:14px;
    color:#888;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------
# LIBREOFFICE PATH
# ---------------------------------------------------
def get_libreoffice_path(): 
    if os.name == "nt":     
        possible_paths = [    
            r"C:\Program Files\LibreOffice\program\soffice.exe",  
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe" 
        ]
        for p in possible_paths:  
            if os.path.exists(p):  
                return p           
    else:  
        return shutil.which("soffice")  
    return None    

libreoffice_path = get_libreoffice_path()  

if libreoffice_path is None: 
    st.error("LibreOffice not found")  
    st.stop()  

# ---------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------
def convert_docx_to_pdf(input_path, output_folder): 
    cmd = [                  
        libreoffice_path,    
        "--headless",       
        "--nologo",          
        "--norestore",       
        "--nofirststartwizard",  
        "--convert-to", "pdf",  
        "--outdir", output_folder,  
        input_path   
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE) 
    
def convert_pdf_to_docx(input_path, output_path):  
    cv = Converter(input_path) 
    cv.convert(output_path, start=0, end=None)  
    cv.close()  

# ---------------------------------------------------
# TABS
# ---------------------------------------------------
tab1, tab2, tab3 = st.tabs([t("title_main"), t("img_title"), t("merge_title")])

# ---------------------------------------------------
# TAB 1 — PDF ↔ Word
# ---------------------------------------------------
with tab1:

    st.title(t("title_main"))

    uploaded = st.file_uploader(
        t("upload"),  
        type=["docx", "pdf"]   
    )

    if uploaded:   

        ext = Path(uploaded.name).suffix.lower()  
        with tempfile.TemporaryDirectory() as tmpdir:   

            input_path = os.path.join(tmpdir, uploaded.name) 

            with open(input_path, "wb") as f: 
                f.write(uploaded.read()) 
                
            progress = st.progress(0) 

            try:  

                # DOCX -> PDF
                if ext == ".docx":  
                    st.info(t("docx_info")) 
                    progress.progress(30)  

                    convert_docx_to_pdf(input_path, tmpdir) 

                    progress.progress(80)

                    output_name = uploaded.name.replace(".docx", ".pdf") 
                    output_path = os.path.join(tmpdir, output_name) 

                    with open(output_path, "rb") as f:  
                        progress.progress(100)   
                        st.success(t("done"))  

                        st.download_button(  
                            t("download_pdf"),  
                            data=f,  
                            file_name=output_name,  
                            mime="application/pdf"  
                        )

                # PDF -> DOCX
                elif ext == ".pdf":  

                    st.info(t("pdf_info"))  
                    progress.progress(30)

                    output_name = uploaded.name.replace(".pdf", ".docx")  
                    output_path = os.path.join(tmpdir, output_name) 

                    convert_pdf_to_docx(input_path, output_path) 

                    progress.progress(100)

                    with open(output_path, "rb") as f: 
                        st.success(t("done"))   

                        st.download_button(  
                            t("download_docx"),  
                            data=f,  
                            file_name=output_name,  
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
                        )

            except Exception as e:  
                st.error(t("error"))  
                st.code(str(e)) 

# ---------------------------------------------------
# TAB 2 — IMAGE → PDF
# ---------------------------------------------------
with tab2:

    st.title(t("img_title"))

    uploaded_images = st.file_uploader(
        t("img_upload"),
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True
    )

    if uploaded_images:
        progress = st.progress(0)
        st.info(t("img_info"))

        try:
            images = []
            for img_file in uploaded_images:
                img = Image.open(img_file).convert("RGB")
                images.append(img)

            progress.progress(60)

            # Creazione PDF temporaneo
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                pdf_path = tmp.name

            if len(images) == 1:
                images[0].save(pdf_path, save_all=True)
            else:
                images[0].save(pdf_path, save_all=True, append_images=images[1:])

            progress.progress(100)
            st.success(t("done"))

            with open(pdf_path, "rb") as f:
                st.download_button(
                    t("download_img_pdf"),
                    data=f,
                    file_name="images.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(t("error"))
            st.code(str(e))

# ---------------------------------------------------
# TAB 3 — MERGE PDF
# ---------------------------------------------------
with tab3:

    st.title(t("merge_title"))

    uploaded_pdfs = st.file_uploader(
        t("merge_upload"),
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_pdfs and len(uploaded_pdfs) >= 2:
        progress = st.progress(0)
        st.info(t("merge_info"))

        try:
            merger = PdfMerger()

            for pdf in uploaded_pdfs:
                merger.append(pdf)

            progress.progress(60)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                merged_path = tmp.name
                merger.write(merged_path)
                merger.close()

            progress.progress(100)
            st.success(t("done"))

            with open(merged_path, "rb") as f:
                st.download_button(
                    t("download_merged_pdf"),
                    data=f,
                    file_name="merged.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(t("error"))
            st.code(str(e))

    elif uploaded_pdfs:
        st.warning("Please upload at least two PDF files." if st.session_state.lang == "en"
                   else "Carica almeno due file PDF.")

# ---------------------------------------------------
# DONATION SECTION
# ---------------------------------------------------
st.markdown("---")

st.markdown(
    f"""
    <div style='text-align:center; margin-top:24px; margin-bottom:8px;'>
        <div style='font-size:16px; font-weight:500;'>
            {t("donation_title")}
        </div>
        <div style='font-size:13px; color:#888; margin-top:6px;'>
            {t("donation_subtitle")}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns([1, 1.2, 1])

with col2:
    st.link_button(
        t("donation_button"),
        "https://www.paypal.com/donate/?hosted_button_id=2YWFSJBJF5WP6",
        use_container_width=True
    )

# ---------------------------------------------------
# FOOTER
# ---------------------------------------------------
st.markdown("---")
st.markdown(
    f"""
    <div class='small'>{t("footer1")}</div>
    <div class='small'>{t("footer2")}</div>
    """,
    unsafe_allow_html=True
)
