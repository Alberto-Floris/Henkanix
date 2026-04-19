import streamlit as st
import tempfile
import os
import subprocess
from pathlib import Path
from pdf2docx import Converter
import shutil

# ---------------------------------------------------
# LANGUAGE SYSTEM
# ---------------------------------------------------
LANG = {
    "en": {
        "title": "Upload a PDF or Word file and convert it in seconds.",
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
        "button_lang": "🌐 Italiano"
    },
    "it": {
        "title": "Carica un file PDF o Word e convertilo in pochi secondi.",
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
        "button_lang": "🌐 English"
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

    button_label = "🌐 Italiano" if current_lang == "en" else "🌐 English"

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
                color: var(--text-color);
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
h1 {text-align:center;}
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
# HEADER
# ---------------------------------------------------
st.title(t("title_main"))

uploaded = st.file_uploader(
    t("upload"),
    type=["docx", "pdf"]
)

# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
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