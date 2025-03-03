import streamlit as st
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.platypus import ListFlowable, ListItem

# Konfigurasi Google Drive (GANTI DENGAN ID FOLDER ANDA)
FOLDER_ID = "1Zvee0AaW9w2p0PLyqQ03bmCdmdt5dX-s"

# Load kredensial dari Streamlit Secrets
creds_dict = json.loads(st.secrets["gdrive_service_account"])
creds = service_account.Credentials.from_service_account_info(creds_dict)
drive_service = build("drive", "v3", credentials=creds)

def export_pdf(data, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    elements = []

    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.alignment = TA_CENTER
    subtitle_style = styles["Heading2"]
    subtitle_style.alignment = TA_CENTER

    answer_style1 = ParagraphStyle("answer_style1", parent=styles["Normal"], alignment=TA_CENTER)
    answer_style2 = ParagraphStyle("answer_style2", parent=styles["Normal"], alignment=TA_JUSTIFY)

    title = Paragraph("SPOT Light", title_style)
    subtitle = Paragraph("Summary of Progress & Objectives Tracker", subtitle_style)
    elements.append(title)
    elements.append(Spacer(1, 2))
    elements.append(subtitle)
    elements.append(Spacer(1, 16))

    for idx, (key, value) in enumerate(data.items()):
        question = Paragraph(f"<b>{key}</b>", styles["Heading2"])
        elements.append(question)
        elements.append(Spacer(1, 6))

        if isinstance(value, str):
            lines = value.split("\n")
            bullet_items = []
            numbered_items = []
            is_numbered = True  # Default as numbered list, change if no numbered items detected

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                match = re.match(r"^(\d+)\.\s(.+)", line)  # Cek angka + titik + spasi
                if match:
                    _, text = match.groups()  # Ambil hanya teks tanpa angka
                    numbered_items.append(ListItem(Paragraph(text, answer_style2)))  # Gunakan numbering otomatis
                elif line.startswith("- "):  # Bulleted list
                    is_numbered = False  # Jika ada bullet, maka bukan numbered list
                    bullet_items.append(ListItem(Paragraph(line[2:], answer_style2)))
                else:
                    is_numbered = False
                    bullet_items.append(ListItem(Paragraph(line, answer_style2)))

            if is_numbered and numbered_items:
                answer = ListFlowable(numbered_items, bulletType="1", leftIndent=20)  # Gunakan numbering otomatis
            else:
                answer = ListFlowable(bullet_items, bulletType="bullet", leftIndent=20)

        else:
            answer_style = answer_style1 if idx <= 4 else answer_style2
            answer = Paragraph(str(value), answer_style)

        elements.append(answer)
        elements.append(Spacer(1, 12))

    doc.build(elements)
    buffer.seek(0)
    return buffer

def upload_to_drive(file_buffer, filename):
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

# Menggunakan Markdown dengan HTML untuk center alignment
st.markdown(
    """
    <h1 style='text-align: center;'>SPOT Light</h1>
    <h3 style='text-align: center;'>Summary of Progress & Objectives Tracker</h3>
    <br>
    """,
    unsafe_allow_html=True
)

# Form Input
with st.form("data_form"):
    # Identitas Tim
    nama_tim = st.selectbox("Nama Tim", ["", 
                                         "[SDGS] Tim Indikator SDGs",
                                         "[IPMPlus] Tim IPM-IPG-IDG-IKG",
                                         "[KMD] Tim Kemiskinan Multidimensi", 
                                         "[ASUS] Tim Analisis Sensus", 
                                         "[CERDAS] Tim Analisis Isu Terkini dan Cerita Data Statistik",
                                         "[SD] Tim Pengembangan Sistem Dinamik",
                                         "[SAE] Tim Small Area Estimation",
                                         "[DS] Tim Data Science",
                                         "[DEV-QA] Tim Pengembangan Penjaminan Kualitas",
                                         "[QG] Tim Quality Gates",
                                         "[SIQAF] Tim Self Assessment dan Pengembangan QAF System",
                                         "[PSS] Tim Pembinaan Statistik Sektoral",
                                         "[REPO] Tim Repository",
                                         "[FMS] Tim Sekretariat Forum Masyarakat Statistik",
                                         "[PVD] Tim Pedoman Validasi Data",
                                         "[SAKIP] Tim SAKIP",
                                         "[GEN-AI] Tim Generative AI",
                                         "[CAN] Tim Pengelola Manajemen Perubahan",
                                         "[MADYA] Tim Penugasan Khusus Tambahan"])
    ketua = st.selectbox("Nama Ketua", ["", 
                                        "Erna Yulianingsih SST, M.Appl.Ecmets", 
                                        "Alvina Clarissa SST", 
                                        "Adi Nugroho, SST",
                                        "Khairunnisah SST, M.S.E", 
                                        "Valent Gigih Saputri SST, M.Ec.Dev.", 
                                        "Nurarifin SST, M.Ec.Dev, M.Ec.", 
                                        "Dhiar Niken Larasati SST, M.E.", 
                                        "Dewi Krismawati SST, M.T.I", 
                                        "Sukmasari Dewanti SST, M.Sc", 
                                        "Reni Amelia, S.S.T., M.Si.", 
                                        "Yohanes Eki Apriliawan SST", 
                                        "Zulfa Hidayah Satria Putri SST, M.Stat.", 
                                        "Mohammad Ammar Alwandi S.Tr.Stat.", 
                                        "Synthia Natalia Kristiani SST", 
                                        "Muhammad Ihsan SST",  
                                        "Putri Wahyu Handayani SST, M.S.E", 
                                        "Dewi Lestari Amaliah SST, M.B.A."])
    coach = st.selectbox("Nama Coach", ["", 
                                        "Dr. Ambar Dwi Santoso S.Si, M.Si", 
                                        "Indah Budiati SST, M.Si", 
                                        "Wisnu Winardi, SST, ME.", 
                                        "Edi Waryono S.Si., M.Kesos.", 
                                        "Mutijo S.Si, M.Si", 
                                        "Usman Bustaman S.Si, M.Sc", 
                                        "Dr. Arham Rivai S.Si, M.Si", 
                                        "Lestyowati Endang Widyantari, S.Si, M.Kesos", 
                                        "Taulina Anggarani S.Si, MA", 
                                        "Dr. Muchammad Romzi"])
    jumlah_anggota = st.number_input("Jumlah Anggota", min_value=1)
    bulan = st.selectbox("Periode Pelaporan (Bulan)", ["", 
                                                       "Januari 2025", 
                                                       "Februari 2025", 
                                                       "Januari-Februari 2025", 
                                                       "Maret 2025", 
                                                       "April 2025", 
                                                       "Mei 2025", 
                                                       "Juni 2025", 
                                                       "Juli 2025", 
                                                       "Agustus 2025", 
                                                       "September 2025", 
                                                       "Oktober 2025", 
                                                       "November 2025", 
                                                       "Desember 2025"])

    # Objective/Goal Tahunan
    objective = st.text_area("Objective/Goal Tahunan")

    # Progress Bulanan vs Target Triwulanan
    st.subheader("Progress Bulanan vs Target Triwulanan")
    progress_bulanan = st.text_area("Progress Bulanan dengan Indikator Pencapaian")
    target_triwulanan = st.text_area("Target Triwulanan dengan Indikator Pencapaian")

    # Hasil Retrospective
    st.subheader("Hasil Retrospektif")
    what_went_well = st.text_area("What went Well?")
    what_can_be_improved = st.text_area("What can be Improved?")
    action_points = st.text_area("Action Points")

    submitted = st.form_submit_button("Simpan Data")

    if submitted:
        st.session_state.data = {
            "Nama Tim": nama_tim,
            "Ketua": ketua,
            "Coach": coach,
            "Jumlah Anggota": jumlah_anggota,
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
        pdf_buffer = export_pdf(st.session_state.data, filename)
        gdrive_link = upload_to_drive(pdf_buffer, filename)
        
        if gdrive_link:
            st.success("PDF berhasil diunggah ke Google Drive!")
            st.markdown(f"[Lihat File di Google Drive]({gdrive_link})")
            st.download_button("Download PDF", pdf_buffer, file_name=filename, mime="application/pdf")
        else:
            st.error("Gagal mengunggah ke Google Drive.")
