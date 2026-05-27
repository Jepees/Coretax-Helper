"""
Coretax Helper — Excel to XML Converter
A cross-platform web app to convert Coretax Excel templates to XML format.
Supports BP21, BPMP, BPA1, and BPPU.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import random
import time

from modules.excel_reader import read_bp21, read_bpmp, read_bpa1, read_bppu
from modules.xml_generator import (
    generate_bp21_xml, generate_bpmp_xml,
    generate_bpa1_xml, generate_bppu_xml,
)
from modules.reference_data import (
    BP21_TAX_OBJECTS,
    BPMP_TAX_OBJECTS,
    BP21_COLUMNS,
    BPMP_COLUMNS,
    BPA1_COLUMNS,
    BPPU_COLUMNS,
)

# =============================================================================
# Page Config
# =============================================================================
st.set_page_config(
    page_title="Coretax Helper",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =============================================================================
# Custom CSS
# =============================================================================
st.markdown("""
<style>
    /* Import Inter font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Global font */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Header styling */
    .main-title {
        background: linear-gradient(135deg, #00D4FF 0%, #0080FF 50%, #7B61FF 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0;
        letter-spacing: -0.02em;
    }

    .subtitle {
        color: #8899AA;
        font-size: 1.05rem;
        margin-top: -0.5rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }

    /* Card header inside container */
    .card-header {
        background: linear-gradient(135deg, #00D4FF 0%, #7B61FF 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 0 0 0.5rem 0;
        font-weight: 700;
        font-size: 1.05rem;
    }

    .card-desc {
        color: #8899AA;
        font-size: 0.82rem;
        line-height: 1.5;
        margin-bottom: 0.6rem;
        min-height: 2.5em;
    }

    .badge-row {
        margin-bottom: 0.8rem;
    }

    /* Force equal card heights — target st.container(border=True) */
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] > div > [data-testid="stVerticalBlockBorderWrapper"] {
        min-height: 260px;
        display: flex;
        flex-direction: column;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] > div > [data-testid="stVerticalBlockBorderWrapper"] > div {
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }

    /* Style the bordered containers as cards */
    [data-testid="stVerticalBlockBorderWrapper"] {
        border-color: rgba(0, 212, 255, 0.15) !important;
        border-radius: 16px !important;
        background: linear-gradient(135deg, rgba(0, 212, 255, 0.04) 0%, rgba(123, 97, 255, 0.04) 100%) !important;
        transition: all 0.3s ease;
    }

    [data-testid="stVerticalBlockBorderWrapper"]:hover {
        border-color: rgba(0, 212, 255, 0.35) !important;
        background: linear-gradient(135deg, rgba(0, 212, 255, 0.08) 0%, rgba(123, 97, 255, 0.08) 100%) !important;
    }

    /* Info card (used in results section) */
    .info-card {
        background: linear-gradient(135deg, rgba(0, 212, 255, 0.08) 0%, rgba(123, 97, 255, 0.08) 100%);
        border: 1px solid rgba(0, 212, 255, 0.2);
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1rem;
    }

    .info-card h4 {
        color: #00D4FF;
        margin: 0 0 0.5rem 0;
        font-weight: 600;
    }

    /* Status badges */
    .badge-success {
        background: rgba(0, 200, 120, 0.15);
        color: #00C878;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
        display: inline-block;
        margin: 0.2rem 0;
    }

    .badge-error {
        background: rgba(255, 75, 75, 0.15);
        color: #FF4B4B;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 500;
        display: inline-block;
        margin: 0.2rem 0;
    }

    .badge-info {
        background: rgba(0, 212, 255, 0.15);
        color: #00D4FF;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 500;
        display: inline-block;
        margin: 0.15rem 0.1rem;
    }

    /* Divider */
    .section-divider {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(0, 212, 255, 0.3), transparent);
        margin: 2rem 0;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #556677;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding: 1rem;
    }

    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00D4FF 0%, #0080FF 100%);
        color: white;
        border: none;
        padding: 0.6rem 2rem;
        font-weight: 600;
        border-radius: 8px;
        transition: all 0.3s ease;
    }

    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #00E5FF 0%, #0090FF 100%);
        box-shadow: 0 4px 20px rgba(0, 212, 255, 0.3);
    }

    /* Data editor styling */
    [data-testid="stDataEditor"] {
        border: 1px solid rgba(0, 212, 255, 0.15);
        border-radius: 8px;
    }

    /* Hide the "+" add-more-files button in file uploader */
    [data-testid="stFileUploader"] [data-testid="stBaseButton-minimal"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# Sidebar — Panduan & Tentang
# =============================================================================
with st.sidebar:
    st.markdown("### 📖 Panduan")
    st.markdown("""
    **Cara Penggunaan:**
    1. Upload/drop file Excel ke card yang sesuai
    2. Periksa data di preview
    3. Klik **Download XML**
    4. Import XML ke Coretax

    **File Excel yang didukung:**
    - `BP21 Excel to XML v.4.xlsx`
    - `BPMP Excel to XML v.3.xlsx`
    - `BPA1 Excel to XML.xlsx`
    - `BPPU Excel to XML v.3.xlsx`
    """)

    st.markdown("---")

    st.markdown("### ℹ️ Tentang")
    st.markdown("""
    **Coretax Helper** menggantikan fitur VBA Export
    di Excel yang hanya tersedia di Windows.

    App ini berjalan di browser — bisa dipakai
    di **Mac**, **Windows**, maupun **Linux**.

    Semua proses dilakukan di sisi server,
    file Excel tidak disimpan.
    """)

# =============================================================================
# Header
# =============================================================================
st.markdown("""
<div class="card" style="padding: 20px 24px; margin-bottom: 24px;">
    <h1 style="margin: 0 0 8px 0; font-size: 2rem; font-weight: 700; color: #00D4FF;">📄 Coretax Helper</h1>
    <p style="margin: 0; color: rgba(250, 250, 250, 0.7); font-size: 1.05rem;">
        Converter Excel to XML untuk Coretax — Upload file Excel ke salah satu tipe di bawah
    </p>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# Card definitions
# =============================================================================
CARD_CONFIG = {
    "BP21": {
        "title": "📋 BP21 — Bukti Potong Pasal 21",
        "desc": (
            "Non-pegawai tetap: tenaga ahli, peserta kegiatan, "
            "peserta perlombaan, pegawai tidak tetap, honorarium, dll."
        ),
        "badges": ["32 kode objek pajak", "TER / PS17 / Harian / Pesangon"],
        "reader": read_bp21,
        "generator": generate_bp21_xml,
        "columns": BP21_COLUMNS,
    },
    "BPMP": {
        "title": "📋 BPMP — Masa Pegawai Tetap",
        "desc": (
            "Payroll/gaji bulanan: pegawai tetap, "
            "penerima pensiun teratur, pegawai di daerah tertentu."
        ),
        "badges": ["3 kode objek pajak", "Tarif TER"],
        "reader": read_bpmp,
        "generator": generate_bpmp_xml,
        "columns": BPMP_COLUMNS,
    },
    "BPA1": {
        "title": "📋 BPA1 — Tahunan A1",
        "desc": (
            "Bukti potong tahunan pegawai tetap: rekap penghasilan, "
            "tunjangan, bonus, zakat, iuran pensiun, dan PPh Pasal 21."
        ),
        "badges": ["3 kode objek pajak", "27 kolom data"],
        "reader": read_bpa1,
        "generator": generate_bpa1_xml,
        "columns": BPA1_COLUMNS,
    },
    "BPPU": {
        "title": "📋 BPPU — PPh Unifikasi",
        "desc": (
            "PPh Pasal 22, 23, 4(2), 15, dan 26: jasa, sewa, "
            "dividen, bunga, royalti, konstruksi, dll."
        ),
        "badges": ["~200 kode objek pajak", "Fixed rate"],
        "reader": read_bppu,
        "generator": generate_bppu_xml,
        "columns": BPPU_COLUMNS,
    },
}

# =============================================================================
# Reset counter & Easter Egg states
# =============================================================================
if "ee_probability" not in st.session_state:
    st.session_state.ee_probability = 10.0

if "success_count" not in st.session_state:
    st.session_state.success_count = 0

if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0
    st.session_state.has_toasted = False
    st.session_state.ee_rolled = False
    st.session_state.ee_active = False
    st.session_state.ee_qtype = None
    st.session_state.ee_no_count = 0
    st.session_state.ee_passed = False

rc = st.session_state.reset_counter

# =============================================================================
# Detect which uploader already has a file (before rendering)
# =============================================================================
locked_type = None
for dt in ["BPMP", "BP21", "BPA1", "BPPU"]:
    key = f"upload_{dt}_{rc}"
    if key in st.session_state and st.session_state[key] is not None:
        locked_type = dt
        break

# =============================================================================
# Reset button — always visible
# =============================================================================
_, reset_col = st.columns([5, 1])
with reset_col:
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.reset_counter += 1
        st.session_state.has_toasted = False
        st.session_state.ee_rolled = False
        st.session_state.ee_active = False
        st.session_state.ee_no_count = 0
        st.session_state.ee_passed = False
        st.session_state.ee_probability += 3.0
        st.rerun()

# =============================================================================
# Upload Cards — 2×2 grid, uploader INSIDE each card
# =============================================================================
uploaded_files = {}

row1_col1, row1_col2 = st.columns(2)
row2_col1, row2_col2 = st.columns(2)

grid = [
    (row1_col1, "BPMP"),
    (row1_col2, "BP21"),
    (row2_col1, "BPA1"),
    (row2_col2, "BPPU"),
]

for col, doc_type_key in grid:
    cfg = CARD_CONFIG[doc_type_key]
    badges_html = " ".join(
        f'<span class="badge-info">{b}</span>' for b in cfg["badges"]
    )
    # Disable ALL uploaders once any file is uploaded
    is_disabled = locked_type is not None

    with col:
        with st.container(border=True):
            st.markdown(
                f'<div style="min-height: 160px; display: flex; flex-direction: column;">'
                f'<h4 class="card-header">{cfg["title"]}</h4>'
                f'<p class="card-desc" style="flex-grow: 1;">{cfg["desc"]}</p>'
                f'<div class="badge-row">{badges_html}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
            uploaded = st.file_uploader(
                f"Upload file Excel {doc_type_key}",
                type=["xlsx"],
                key=f"upload_{doc_type_key}_{rc}",
                label_visibility="collapsed",
                disabled=is_disabled,
            )
            if uploaded:
                uploaded_files[doc_type_key] = uploaded

# =============================================================================
# Determine active type
# =============================================================================
active_type = locked_type
active_file = uploaded_files.get(active_type) if active_type else None

# =============================================================================
# Process & show results
# =============================================================================
if active_type and active_file:
    cfg = CARD_CONFIG[active_type]
    file_bytes = active_file.read()

    loading_msgs = [
        "Bentar ya sayang, datanya lagi diproses...",
        "Semangat terus kerjanya cantikk! ❤️",
        "Tunggu sebentar ya, lagi disiapin file-nya...",
        "Kerja yang pinter ya, jangan terlalu capek!",
        "Sabar ya, sistemnya lagi muter otak buat bantu kamu...",
    ]
    
    with st.spinner(random.choice(loading_msgs)):
        if not st.session_state.has_toasted:
            time.sleep(random.uniform(2.0, 2.5))
        result = cfg["reader"](file_bytes)

    tin = result["tin"]
    rows = result["rows"]
    errors = result["errors"]

    # =========================================================================
    # Easter Egg Toast (Hanya muncul sekali per upload)
    # =========================================================================
    if not st.session_state.has_toasted and not errors:
        st.session_state.success_count += 1
        n = st.session_state.success_count
        
        if n % 5 == 0 and n > 0:
            toast_msgs = [
                f"Hebat! Kamu udah nyelesaiin {n} laporan nih! Istirahat bentar dan jangan lupa minum air yaa 💧",
                f"Wow, {n} file selesai! Matanya diistirahatin dulu ya cantik 👀",
                f"Udah {n} laporan nih, bangga banget! Tarik nafas panjang dulu yuk 🧘‍♀️",
                f"Progress luar biasa: {n} file beres! Jangan lupa senyum hari ini 😊",
            ]
        else:
            toast_msgs = [
                "Sukses! Semangat terus kerjanya ya! ✨",
                "Jangan lupa minum air putih ya 💧",
                "Udah makan belum? Jangan sampai telat makan 🍚",
                "Kerjaan emang banyak, tapi kamu pasti bisa. Proud of you!",
                "Ada yang lagi kangen kamu nih di sini... 🥺",
                "Jangan lupa istirahat kalau matanya udah capek ya.",
            ]
        st.toast(random.choice(toast_msgs), icon="✨")
        st.session_state.has_toasted = True

    # =========================================================================
    # Step 1: Validation & Info
    # =========================================================================
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown(f"### ① Hasil Pembacaan File — {active_type}")

    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
        if tin and len(tin) == 16:
            st.markdown(
                f'<div class="info-card">'
                f'<h4>NPWP Pemotong</h4>'
                f'<p style="font-size: 1.1rem; font-weight: 600; color: #00C878; margin: 0;">✅ {tin}</p>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="info-card">'
                f'<h4>NPWP Pemotong</h4>'
                f'<p style="font-size: 1.1rem; font-weight: 600; color: #FF4B4B; margin: 0;">❌ {tin or "Tidak ditemukan"}</p>'
                f'</div>',
                unsafe_allow_html=True,
            )

    with col_r2:
        st.markdown(
            f'<div class="info-card">'
            f'<h4>Jumlah Data</h4>'
            f'<p style="font-size: 1.1rem; font-weight: 600; color: #00D4FF; margin: 0;">📊 {len(rows)} baris</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col_r3:
        st.markdown(
            f'<div class="info-card">'
            f'<h4>Tipe</h4>'
            f'<p style="font-size: 1.1rem; font-weight: 600; color: #00D4FF; margin: 0;">📄 {active_type}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Show errors/warnings
    if errors:
        with st.expander(f"⚠️ {len(errors)} peringatan ditemukan", expanded=True):
            for err in errors:
                st.warning(err)

    # =========================================================================
    # Step 2: Data Preview (editable)
    # =========================================================================
    if rows:
        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        st.markdown("### ② Preview Data")
        st.caption("Anda dapat mengedit data di tabel sebelum generate XML.")

        df = pd.DataFrame(rows)

        columns = cfg["columns"]
        column_config = {}
        column_order = []

        for label, key, col_type in columns:
            column_order.append(key)
            if col_type == "int":
                column_config[key] = st.column_config.NumberColumn(
                    label, min_value=0, format="%d"
                )
            elif col_type == "float":
                column_config[key] = st.column_config.NumberColumn(
                    label, format="%.2f"
                )
            elif col_type == "date":
                column_config[key] = st.column_config.TextColumn(label)
            elif col_type == "select":
                column_config[key] = st.column_config.TextColumn(label)
            else:
                column_config[key] = st.column_config.TextColumn(label)

        edited_df = st.data_editor(
            df,
            column_config=column_config,
            column_order=column_order,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
        )

        edited_rows = edited_df.to_dict("records")

        # =====================================================================
        # Step 3: Generate & Preview XML
        # =====================================================================
        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        st.markdown("### ③ Download & Preview XML")

        xml_output = cfg["generator"](tin, edited_rows)

        col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
        with col_dl2:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{active_type.lower()}_{timestamp}.xml"

            # Roll easter egg based on dynamic probability
            if not st.session_state.ee_rolled:
                if random.random() < (st.session_state.ee_probability / 100.0):
                    st.session_state.ee_active = True
                    st.session_state.ee_qtype = random.choice(["A", "B"])
                    # Reset probability after trigger
                    st.session_state.ee_probability = 10.0
                st.session_state.ee_rolled = True

            if st.session_state.ee_active and not st.session_state.ee_passed:
                # Show the easter egg question dynamically based on no_count
                if st.session_state.ee_qtype == "A":
                    if st.session_state.ee_no_count == 0:
                        q_text = "Sebelum download, jawab dulu: Hari ini kangen pembuat app ini nggak?"
                    elif st.session_state.ee_no_count == 1:
                        q_text = "Yang bener? Coba pikir-pikir lagi deh..."
                    else:
                        q_text = "Yang bener, masa nggak kangen sih, sedih nih 😢"
                else:
                    if st.session_state.ee_no_count == 0:
                        q_text = "Sebelum download, jawab dulu: masih mau balikan nggak?"
                    elif st.session_state.ee_no_count == 1:
                        q_text = "Masa sih? Coba jujur sama diri sendiri..."
                    else:
                        q_text = "Jangan bohong dong, ayo ngaku aja! 🥺"
                
                st.markdown(f"<div style='text-align: center; color: #FF4B4B; font-weight: bold; margin-bottom: 10px;'>{q_text}</div>", unsafe_allow_html=True)
                
                ecol1, ecol2 = st.columns(2)
                
                if st.session_state.ee_no_count >= 2:
                    # Buttons change completely
                    if st.session_state.ee_qtype == "A":
                        b1_text, b2_text = "Iyaaaaaa kangen", "Kangen bangettttt"
                    else:
                        b1_text, b2_text = "Iya masih mauuuu balikan bangett", "Iya mauuuuu banget"
                        
                    with ecol1:
                        if st.button(b1_text, use_container_width=True, type="primary"):
                            st.session_state.ee_passed = True
                            st.balloons()
                            st.rerun()
                    with ecol2:
                        if st.button(b2_text, use_container_width=True, type="primary"):
                            st.session_state.ee_passed = True
                            st.balloons()
                            st.rerun()
                else:
                    # Initial buttons
                    with ecol1:
                        if st.button("Iya", use_container_width=True, type="primary"):
                            st.session_state.ee_passed = True
                            st.balloons()
                            st.rerun()
                    with ecol2:
                        if st.button("Enggak", use_container_width=True):
                            st.session_state.ee_no_count += 1
                            st.rerun()
            else:
                # Normal download button
                if st.session_state.ee_active and st.session_state.ee_passed:
                    st.markdown("<div style='text-align: center; color: #00D4FF; margin-bottom: 8px; font-size: 0.95rem;'>kalo kangen/pengen balikan mah gausah gengsi-gengsi kali, ngomong aja 😏😌</div>", unsafe_allow_html=True)

                st.download_button(
                    label=f"📥 Download XML — {filename}",
                    data=xml_output.encode("utf-8"),
                    file_name=filename,
                    mime="application/xml",
                    use_container_width=True,
                )
                st.caption("<div style='text-align: center;'>File XML siap di-import ke Coretax melalui menu eBupot.</div>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        with st.expander("🔍 Preview XML", expanded=False):
            st.code(xml_output, language="xml")

# =============================================================================
# Footer
# =============================================================================
st.markdown(
    '<div class="footer">'
    'Coretax Excel to XML Helper — Cross-platform Excel to XML Converter<br>'
    'Tidak terafiliasi dengan Direktorat Jenderal Pajak<br>'
    'Made By Jeepes for JP'
    '</div>',
    unsafe_allow_html=True,
)
