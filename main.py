import streamlit as st
import tempfile
import os
import subprocess
from pathlib import Path
from pdf2docx import Converter
import shutil

# ---------------------------------------------------
# CONFIG PAGINA
# ---------------------------------------------------
st.set_page_config(
    page_title="Henkanix",
    page_icon="✨",
    layout="centered"
)

# ---------------------------------------------------
# TITOLO PRINCIPALE
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
    """
      <div style='text-align:center; font-size:20px;
                color: var(--text-color);
                margin-top:16px; line-height:1.8;
                max-width:780px; margin-left:auto; margin-right:auto;
                font-weight:500;'>
        Carica un file <b>PDF o Word</b> e convertilo in pochi secondi.<br>
        Nessuna installazione richiesta, tutto funziona direttamente dal browser.
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
# LIBREOFFICE PATH (FUNZIONA OVUNQUE)
# ---------------------------------------------------
def get_libreoffice_path():
    if os.name == "nt":  # Windows
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for p in possible_paths:
            if os.path.exists(p):
                return p
    else:  # Linux (Streamlit Cloud)
        return shutil.which("soffice")

    return None


libreoffice_path = get_libreoffice_path()

if libreoffice_path is None:
    st.error("LibreOffice non trovato nel sistema")
    st.stop()

# ---------------------------------------------------
# FUNZIONI
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
st.title("PDF ↔ Word")

uploaded = st.file_uploader(
    "Carica file 📄",
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

            # -------------------------------------
            # DOCX -> PDF
            # -------------------------------------
            if ext == ".docx":

                st.info("Conversione Word → PDF in corso...")
                progress.progress(30)

                convert_docx_to_pdf(input_path, tmpdir)

                progress.progress(80)

                output_name = uploaded.name.replace(".docx", ".pdf")
                output_path = os.path.join(tmpdir, output_name)

                with open(output_path, "rb") as f:
                    progress.progress(100)
                    st.success("Conversione completata!")

                    st.download_button(
                        "⬇️ Scarica PDF",
                        data=f,
                        file_name=output_name,
                        mime="application/pdf"
                    )

            # -------------------------------------
            # PDF -> DOCX
            # -------------------------------------
            elif ext == ".pdf":

                st.info("Conversione PDF → Word in corso...")
                progress.progress(30)

                output_name = uploaded.name.replace(".pdf", ".docx")
                output_path = os.path.join(tmpdir, output_name)

                convert_pdf_to_docx(input_path, output_path)

                progress.progress(100)

                with open(output_path, "rb") as f:
                    st.success("Conversione completata!")

                    st.download_button(
                        "⬇️ Scarica Word",
                        data=f,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        except Exception as e:
            st.error("Errore durante la conversione.")
            st.code(str(e))

# ---------------------------------------------------
# FOOTER
# ---------------------------------------------------
st.markdown("---")
st.markdown(
    """
    <div class='small'>Powered by Streamlit + LibreOffice + Python</div>
    <div class='small'>Created by Alberto Floris</div>
    """,
    unsafe_allow_html=True
)