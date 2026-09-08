import streamlit as st
import tempfile
import os
import subprocess
import zipfile
import io
import numpy as np
from pathlib import Path
from pdf2docx import Converter
import shutil
from PIL import Image
from PyPDF2 import PdfMerger
import fitz  # PyMuPDF - per PDF -> Immagini e Firma PDF
from streamlit_drawable_canvas import st_canvas

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

        # IMAGE <-> PDF
        "img_title": "Images ↔ PDF",
        "direction_label": "Choose conversion direction",
        "direction_img_to_pdf": "Image(s) → PDF",
        "direction_pdf_to_img": "PDF → Image(s)",
        "img_upload": "Upload one or more images 🖼️",
        "img_info": "Image → PDF conversion in progress...",
        "download_img_pdf": "⬇️ Download PDF",
        "pdf_to_img_upload": "Upload a PDF file 📄",
        "pdf_to_img_info": "PDF → Image conversion in progress...",
        "pdf_to_img_format": "Output format",
        "pdf_to_img_dpi": "Quality (DPI)",
        "download_single_image": "⬇️ Download image",
        "download_zip_images": "⬇️ Download images (ZIP)",

        # MERGE PDF
        "merge_title": "Merge multiple PDFs",
        "merge_upload": "Upload two or more PDF files 📚",
        "merge_info": "Merging PDF files...",
        "download_merged_pdf": "⬇️ Download merged PDF",

        # SIGN PDF
        "sign_title": "Sign PDF",
        "sign_source_label": "How do you want to add your signature?",
        "sign_draw": "Draw signature",
        "sign_upload": "Upload signature image",
        "sign_draw_instructions": "Draw your signature below with your mouse or finger.",
        "sign_upload_label": "Upload a signature image (ideally transparent PNG) 🖊️",
        "sign_pdf_upload": "Upload the PDF to sign 📄",
        "sign_page_number": "Page to sign",
        "sign_pos_x": "Horizontal position (%)",
        "sign_pos_y": "Vertical position (%)",
        "sign_width": "Signature width (% of page)",
        "sign_preview_caption": "Preview",
        "sign_download": "⬇️ Download signed PDF",
        "sign_missing_signature": "Draw or upload a signature first.",
        "sign_disclaimer": "This is a visual signature (an image stamped onto the page), not a legally certified digital signature.",
        "sign_make_transparent": "Remove white background from signature",
        "sign_transparency_sensitivity": "Background sensitivity",

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

        # IMAGE <-> PDF
        "img_title": "Immagini ↔ PDF",
        "direction_label": "Scegli la direzione della conversione",
        "direction_img_to_pdf": "Immagine/i → PDF",
        "direction_pdf_to_img": "PDF → Immagine/i",
        "img_upload": "Carica una o più immagini 🖼️",
        "img_info": "Conversione Immagini → PDF in corso...",
        "download_img_pdf": "⬇️ Scarica PDF",
        "pdf_to_img_upload": "Carica un file PDF 📄",
        "pdf_to_img_info": "Conversione PDF → Immagini in corso...",
        "pdf_to_img_format": "Formato di output",
        "pdf_to_img_dpi": "Qualità (DPI)",
        "download_single_image": "⬇️ Scarica immagine",
        "download_zip_images": "⬇️ Scarica immagini (ZIP)",

        # MERGE PDF
        "merge_title": "Unisci più PDF",
        "merge_upload": "Carica due o più PDF 📚",
        "merge_info": "Unione dei PDF in corso...",
        "download_merged_pdf": "⬇️ Scarica PDF unito",

        # FIRMA PDF
        "sign_title": "Firma PDF",
        "sign_source_label": "Come vuoi aggiungere la tua firma?",
        "sign_draw": "Disegna la firma",
        "sign_upload": "Carica immagine firma",
        "sign_draw_instructions": "Disegna la tua firma qui sotto con mouse o dito.",
        "sign_upload_label": "Carica un'immagine della firma (idealmente PNG trasparente) 🖊️",
        "sign_pdf_upload": "Carica il PDF da firmare 📄",
        "sign_page_number": "Pagina da firmare",
        "sign_pos_x": "Posizione orizzontale (%)",
        "sign_pos_y": "Posizione verticale (%)",
        "sign_width": "Larghezza firma (% della pagina)",
        "sign_preview_caption": "Anteprima",
        "sign_download": "⬇️ Scarica PDF firmato",
        "sign_missing_signature": "Disegna o carica prima una firma.",
        "sign_disclaimer": "Questa è una firma visiva (un'immagine sovrapposta alla pagina), non una firma digitale con valore legale certificato.",
        "sign_make_transparent": "Rendi trasparente lo sfondo della firma",
        "sign_transparency_sensitivity": "Sensibilità sfondo",

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

def convert_pdf_to_images(pdf_bytes, dpi=200, img_format="PNG"):
    """
    Renderizza ogni pagina di un PDF in un'immagine.
    Ritorna una lista di tuple (nome_file, bytes_immagine).
    """
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    zoom = dpi / 72  # 72 DPI è la risoluzione base di PDF/fitz
    matrix = fitz.Matrix(zoom, zoom)

    ext = img_format.lower()
    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        pix = page.get_pixmap(matrix=matrix)
        img_bytes = pix.tobytes(ext if ext != "jpg" else "jpeg")
        filename = f"page_{page_index + 1}.{ext}"
        images.append((filename, img_bytes))

    doc.close()
    return images

def make_signature_transparent(img: Image.Image, sensitivity: int = 200) -> Image.Image:
    """
    Rende trasparente lo sfondo chiaro di un'immagine di firma (es. foto su carta bianca).
    sensitivity: soglia di luminosità (0-255). Pixel più chiari della soglia diventano
    progressivamente trasparenti; pixel scuri (l'inchiostro) restano opachi.
    """
    img = img.convert("RGBA")
    data = np.array(img).astype(np.float32)

    r, g, b = data[:, :, 0], data[:, :, 1], data[:, :, 2]
    luminance = 0.299 * r + 0.587 * g + 0.114 * b

    alpha = np.clip((sensitivity - luminance) * (255.0 / max(sensitivity, 1)), 0, 255)

    out = data.astype(np.uint8).copy()
    out[:, :, 3] = alpha.astype(np.uint8)
    return Image.fromarray(out, "RGBA")

# ---------------------------------------------------
# TABS
# ---------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    t("title_main"), t("img_title"), t("merge_title"), t("sign_title")
])

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
# TAB 2 — IMAGES ↔ PDF
# ---------------------------------------------------
with tab2:

    st.title(t("img_title"))

    direction = st.radio(
        t("direction_label"),
        [t("direction_img_to_pdf"), t("direction_pdf_to_img")],
        horizontal=True
    )

    # -------------------------------------------
    # Immagine/i -> PDF
    # -------------------------------------------
    if direction == t("direction_img_to_pdf"):

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

    # -------------------------------------------
    # PDF -> Immagine/i
    # -------------------------------------------
    else:

        uploaded_pdf = st.file_uploader(
            t("pdf_to_img_upload"),
            type=["pdf"]
        )

        col_a, col_b = st.columns(2)
        with col_a:
            img_format = st.selectbox(t("pdf_to_img_format"), ["PNG", "JPG"])
        with col_b:
            dpi = st.slider(t("pdf_to_img_dpi"), min_value=72, max_value=300, value=200, step=1)

        if uploaded_pdf:
            progress = st.progress(0)
            st.info(t("pdf_to_img_info"))

            try:
                pdf_bytes = uploaded_pdf.read()
                progress.progress(30)

                images = convert_pdf_to_images(pdf_bytes, dpi=dpi, img_format=img_format)
                progress.progress(80)

                if len(images) == 1:
                    filename, img_bytes = images[0]
                    mime = "image/png" if img_format == "PNG" else "image/jpeg"

                    progress.progress(100)
                    st.success(t("done"))

                    st.download_button(
                        t("download_single_image"),
                        data=img_bytes,
                        file_name=filename,
                        mime=mime
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for filename, img_bytes in images:
                            zf.writestr(filename, img_bytes)
                    zip_buffer.seek(0)

                    progress.progress(100)
                    st.success(t("done"))

                    st.download_button(
                        t("download_zip_images"),
                        data=zip_buffer,
                        file_name="images.zip",
                        mime="application/zip"
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
# TAB 4 — FIRMA PDF
# ---------------------------------------------------
with tab4:

    st.title(t("sign_title"))
    st.caption(t("sign_disclaimer"))

    sign_source = st.radio(
        t("sign_source_label"),
        [t("sign_draw"), t("sign_upload")],
        horizontal=True
    )

    signature_img = None

    # --- Opzione A: disegna la firma ---
    if sign_source == t("sign_draw"):
        st.caption(t("sign_draw_instructions"))
        canvas_result = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="rgba(255, 255, 255, 0)",
            height=150,
            width=450,
            drawing_mode="freedraw",
            key="signature_canvas",
            return_image_data=True,
        )
        if canvas_result.image_data is not None and canvas_result.image_data[:, :, 3].sum() > 0:
            signature_img = Image.fromarray(canvas_result.image_data.astype("uint8"), "RGBA")

    # --- Opzione B: carica immagine firma ---
    else:
        sig_file = st.file_uploader(t("sign_upload_label"), type=["png", "jpg", "jpeg"])
        if sig_file:
            raw_img = Image.open(io.BytesIO(sig_file.getvalue())).convert("RGBA")

            remove_bg = st.checkbox(t("sign_make_transparent"), value=True)

            if remove_bg:
                sensitivity = st.slider(t("sign_transparency_sensitivity"), 100, 250, 200)
                signature_img = make_signature_transparent(raw_img, sensitivity=sensitivity)
                st.image(signature_img, caption=t("sign_preview_caption"), width=300)
            else:
                signature_img = raw_img

    st.markdown("---")
    pdf_to_sign = st.file_uploader(t("sign_pdf_upload"), type=["pdf"])

    if pdf_to_sign and not signature_img:
        st.info(t("sign_missing_signature"))

    if pdf_to_sign and signature_img:
        try:
            pdf_bytes = pdf_to_sign.getvalue()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            n_pages = len(doc)

            page_number = st.number_input(
                t("sign_page_number"), min_value=1, max_value=n_pages, value=n_pages
            )

            col1, col2, col3 = st.columns(3)
            with col1:
                pos_x = st.slider(t("sign_pos_x"), 0, 100, 65)
            with col2:
                pos_y = st.slider(t("sign_pos_y"), 0, 100, 85)
            with col3:
                width_pct = st.slider(t("sign_width"), 5, 60, 25)

            page = doc.load_page(page_number - 1)
            page_rect = page.rect

            sig_w_pt = page_rect.width * (width_pct / 100)
            aspect = signature_img.height / signature_img.width
            sig_h_pt = sig_w_pt * aspect

            x0 = min(page_rect.width * (pos_x / 100), page_rect.width - sig_w_pt)
            y0 = min(page_rect.height * (pos_y / 100), page_rect.height - sig_h_pt)
            rect = fitz.Rect(x0, y0, x0 + sig_w_pt, y0 + sig_h_pt)

            sig_buffer = io.BytesIO()
            signature_img.save(sig_buffer, format="PNG")
            page.insert_image(rect, stream=sig_buffer.getvalue())

            # Anteprima della pagina firmata
            pix = page.get_pixmap(matrix=fitz.Matrix(1.3, 1.3))
            st.image(pix.tobytes("png"), caption=t("sign_preview_caption"), use_container_width=True)

            out_buffer = io.BytesIO()
            doc.save(out_buffer)
            doc.close()
            out_buffer.seek(0)

            st.download_button(
                t("sign_download"),
                data=out_buffer,
                file_name="signed_" + pdf_to_sign.name,
                mime="application/pdf"
            )

        except Exception as e:
            st.error(t("error"))
            st.code(str(e))

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