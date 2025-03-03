import streamlit as st
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, ListFlowable, ListItem
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import io
import json
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors

# Konfigurasi Google Drive
FOLDER_ID = "1Zvee0AaW9w2p0PLyqQ03bmCdmdt5dX-s"

# Load kredensial dari Streamlit Secrets
creds_dict = json.loads(st.secrets["gdrive_service_account"])
creds = service_account.Credentials.from_service_account_info(creds_dict)
drive_service = build("drive", "v3", credentials=creds)

# Registrasi font
pdfmetrics.registerFont(TTFont("Lato-Regular", "Lato-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Lato-Bold", "Lato-Bold.ttf"))

# Warna
TURQUOISE = colors.HexColor("#0ba8ed")
DARK_BLUE = colors.HexColor("#041c54")
LIGHT_GRAY = colors.HexColor("#DDDDDD")

# **Fungsi untuk Memproses Bullet List**
def process_text(text, answer_style):
    lines = text.split("\n")
    bullet_items = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("- "):  # Bullet list
            bullet_items.append(ListItem(Paragraph(line[2:], answer_style)))
        else:
            bullet_items.append(ListItem(Paragraph(f"<b>{line}</b>", answer_style)))  # Heading tetap bold

    return ListFlowable(
        bullet_items,
        bulletType="bullet",
        leftIndent=15,
        bulletFontName="Lato-Regular",
        bulletFontSize=12,
        bulletIndent=5
    )

# **Fungsi untuk Ekspor PDF**
def export_pdf(data, filename, logo_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=60, leftMargin=60, topMargin=60, bottomMargin=60)
    elements = []

    styles = getSampleStyleSheet()

    # Gaya teks
    title_style = ParagraphStyle("Title", parent=styles["Title"], fontName="Lato-Bold", fontSize=26, textColor=DARK_BLUE, alignment=TA_CENTER)
    subtitle_style = ParagraphStyle("Subtitle", parent=styles["Heading2"], fontName="Lato-Bold", fontSize=18, textColor=TURQUOISE, alignment=TA_CENTER)
    answer_style = ParagraphStyle("answer_style", parent=styles["Normal"], fontName="Lato-Regular", fontSize=12, alignment=TA_JUSTIFY, leading=18)

    # Header: Logo & Judul
    if logo_path:
        logo = Image(logo_path, width=102.3, height=43.1)
        elements.append(logo)
        elements.append(Spacer(1, 20))
      
    elements.append(Paragraph("SPOT Light", title_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("Summary of Progress & Objectives Tracker", subtitle_style))
    elements.append(Spacer(1, 24))

    # **Loop untuk Menambahkan Data**
    for key, value in data.items():
        elements.append(Paragraph(f"<b>{key}</b>", ParagraphStyle("Question", parent=styles["Heading2"], fontName="Lato-Bold", alignment=TA_CENTER, textColor=DARK_BLUE, leading=18)))
        elements.append(Spacer(1, 6))

        if isinstance(value, str):
            elements.append(process_text(value, answer_style))
        else:
            elements.append(Paragraph(str(value), answer_style))

        elements.append(Spacer(1, 12))

    # **Footer**
    elements.append(Spacer(1, 30))
    elements.append(Table([[""]], colWidths=[500], rowHeights=[1], style=[("BACKGROUND", (0, 0), (-1, -1), LIGHT_GRAY)]))
    elements.append(Spacer(1, 6))
    footer_text = Paragraph("Laporan Tim Direktorat Analisis dan Pengembangan Statistik", ParagraphStyle("Footer", parent=styles["Normal"], fontSize=12, textColor=colors.grey, alignment=TA_CENTER))
    elements.append(footer_text)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# **Fungsi Upload ke Google Drive**
def upload_to_drive(file_buffer, filename):
    query = f"name='{filename}' and '{FOLDER_ID}' in parents and trashed=false"
    existing_files = drive_service.files().list(q=query, fields="files(id)").execute()

    for file in existing_files.get("files", []):
        drive_service.files().delete(fileId=file["id"]).execute()

    file_metadata = {
        "name": filename,
        "parents": [FOLDER_ID]
    }
    
    media = MediaIoBaseUpload(io.BytesIO(file_buffer.getvalue()), mimetype="application/pdf")

    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    return f"https://drive.google.com/file/d/{file['id']}/view"

# **Streamlit UI**
st.markdown(
    """
    <h1 style='text-align: center;'>SPOT Light</h1>
    <h3 style='text-align: center;'>Summary of Progress & Objectives Tracker</h3>
    <br>
    """,
    unsafe_allow_html=True
)

with st.form("data_form"):
    nama_tim = st.text_input("Nama Tim")
    bulan = st.text_input("Periode Pelaporan (Bulan)")
    objective = st.text_area("Objective/Goal Tahunan")
    progress_bulanan = st.text_area("Progress Bulanan")
    target_triwulanan = st.text_area("Target Triwulanan")
    what_went_well = st.text_area("What went Well?")
    what_can_be_improved = st.text_area("What can be Improved?")
    action_points = st.text_area("Action Points")

    submitted = st.form_submit_button("Simpan Data")

    if submitted:
        st.session_state.data = {
            "Nama Tim": nama_tim,
            "Periode Pelaporan": bulan,
            "Objective/Goal Tahunan": objective,
            "Progress Bulanan": progress_bulanan,
            "Target Triwulanan": target_triwulanan,
            "What went Well?": what_went_well,
            "What can be Improved?": what_can_be_improved,
            "Action Points": action_points
        }
        st.success("Data berhasil disimpan!")

if st.session_state.get("data"):
    st.write("### Data Tersimpan")
    st.json(st.session_state.data)

if st.button("Ekspor & Unggah ke Google Drive"):
    if not nama_tim:
        st.warning("Harap isi Nama Tim terlebih dahulu.")
    else:
        filename = f"{nama_tim}_{bulan}.pdf"
        pdf_buffer = export_pdf(st.session_state.data, filename, "logo.png")
        gdrive_link = upload_to_drive(pdf_buffer, filename)
        
        if gdrive_link:
            st.success("PDF berhasil diunggah ke Google Drive!")
            st.markdown(f"[Lihat File di Google Drive]({gdrive_link})")
            st.download_button("Download PDF", pdf_buffer, file_name=filename, mime="application/pdf")
        else:
            st.error("Gagal mengunggah ke Google Drive.")
