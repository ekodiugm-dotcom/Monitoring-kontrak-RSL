
import pandas as pd
import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="Dashboard Monitoring Progress DIV RSL",
    page_icon="📊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "MONITORING PROGRESS DIV RSL.xlsx"
SHEETS = ["MON.PROGRESS-232425", "MON.PROGRESS 2026"]


@st.cache_data
def load_data():
    frames = []

    for sheet in SHEETS:
        df = pd.read_excel(DATA_FILE, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]

        required = ["NOMOR KONTRAK/SPMK", "NAMA PEKERJAAN", "COD", "CHARGECODE"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"{sheet}: missing columns {missing}")

        valid = (
            df["NOMOR KONTRAK/SPMK"].notna()
            & df["NAMA PEKERJAAN"].notna()
            & df["NOMOR KONTRAK/SPMK"].astype(str).str.strip().ne("")
            & df["NAMA PEKERJAAN"].astype(str).str.strip().ne("")
        )

        df = df.loc[valid].copy()
        df["Contract"] = df["NOMOR KONTRAK/SPMK"].astype(str).str.strip()
        df["Scope"] = df["NAMA PEKERJAAN"].astype(str).str.strip()
        # COD is taken from the original COD field in the source workbook.
        # Blank COD values are preserved as blank; they are not inferred.
        df["COD"] = df["COD"].fillna("").astype(str).str.strip()
        df["CHARGECODE"] = df["CHARGECODE"].fillna("").astype(str).str.strip()
        df["Unique Code"] = (
            df["Contract"] + " | " + df["Scope"] + " | " + df["COD"]
        )
        df["Source"] = sheet
        frames.append(df)

    data = pd.concat(frames, ignore_index=True, sort=False)

    # The agreed unique identity is Contract + Scope + COD.
    if data["Unique Code"].duplicated().any():
        raise ValueError("Duplicate Unique Code detected in source data.")

    return data


def clean_text(series):
    return series.fillna("").astype(str).str.strip()


def status_counts(df):
    status = clean_text(df["STATUS PEKERJAAN"])
    return status.replace("", "Data Tidak Tersedia").value_counts()


try:
    df = load_data()
except Exception as exc:
    st.error(f"Data loading error: {exc}")
    st.stop()

st.title("Dashboard Monitoring Progress DIV RSL")
st.caption("Prototype V2 — Executive Summary & Monitoring Detail")

# Monitoring period / source information
st.info(
    f"Population: **{len(df):,} Unique Code**  |  "
    f"Source: **MON.PROGRESS-232425 + MON.PROGRESS 2026**"
)

# -------------------------
# Filters
# -------------------------
st.sidebar.header("Filter Monitoring")

def options(col):
    if col not in df.columns:
        return []
    values = clean_text(df[col])
    return sorted([x for x in values.unique() if x])

filter_fields = [
    ("BIDANG", "Bidang"),
    ("LINGKUP", "Lingkup"),
    ("STATUS PEKERJAAN", "Status Pekerjaan"),
    ("PTL", "PTL"),
    ("PA", "PA"),
    ("Contract", "Contract"),
]

filtered = df.copy()

for col, label in filter_fields:
    if col in filtered.columns:
        opts = options(col)
        selected = st.sidebar.multiselect(label, opts)
        if selected:
            filtered = filtered[clean_text(filtered[col]).isin(selected)]

search = st.sidebar.text_input("Search Contract / Scope / COD")
if search:
    q = search.strip().lower()
    mask = (
        clean_text(filtered["Contract"]).str.lower().str.contains(q, na=False)
        | clean_text(filtered["Scope"]).str.lower().str.contains(q, na=False)
        | clean_text(filtered["COD"]).str.lower().str.contains(q, na=False)
        | clean_text(filtered["Unique Code"]).str.lower().str.contains(q, na=False)
    )
    filtered = filtered[mask]

st.sidebar.caption(f"Filtered population: {len(filtered):,} Unique Code")

# -------------------------
# Executive Summary
# -------------------------
st.header("Executive Summary")

counts = status_counts(filtered)

def n(status):
    return int(counts.get(status, 0))

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Unique Code", f"{len(filtered):,}")
c2.metric("Selesai", f"{n('Selesai'):,}")
c3.metric("Ongoing", f"{n('Ongoing'):,}")
c4.metric("Belum Dilaksanakan", f"{n('Belum Dilaksanakan'):,}")
c5.metric("Batal / Kerja Kurang", f"{n('Batal/Kerja Kurang'):,}")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Status Pekerjaan")
    status_df = counts.rename("Jumlah").to_frame()
    status_df.index.name = "Status"
    st.bar_chart(status_df)

with col2:
    st.subheader("Monitoring by Bidang")
    if "BIDANG" in filtered.columns:
        bidang = (
            clean_text(filtered["BIDANG"])
            .replace("", "Data Tidak Tersedia")
            .value_counts()
            .rename("Jumlah")
            .to_frame()
        )
        bidang.index.name = "Bidang"
        st.bar_chart(bidang)
    else:
        st.info("Kolom BIDANG tidak tersedia.")

# Executive factual attention
st.subheader("Executive Attention")
att1, att2, att3 = st.columns(3)

att1.metric("Ongoing", f"{n('Ongoing'):,}")
att2.metric("Belum Dilaksanakan", f"{n('Belum Dilaksanakan'):,}")
att3.metric("Batal / Kerja Kurang", f"{n('Batal/Kerja Kurang'):,}")

# -------------------------
# Monitoring detail
# -------------------------
st.divider()
st.header("Monitoring Detail")

detail_cols = [
    "Unique Code",
    "Contract",
    "Scope",
    "COD",
    "CHARGECODE",
    "BIDANG",
    "LINGKUP",
    "STATUS PEKERJAAN",
    "PTL",
    "PA",
    "PROGRESS BAYAR",
    "Source",
]
detail_cols = [c for c in detail_cols if c in filtered.columns]

detail = filtered[detail_cols].copy()

# Export the currently filtered monitoring population.
from io import BytesIO

def export_to_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Filtered Data")
    return output.getvalue()

export_col1, export_col2 = st.columns([1, 4])
with export_col1:
    st.download_button(
        label="📥 Export Filtered Data",
        data=export_to_excel(detail),
        file_name="Monitoring_Progress_DIV_RSL_Filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with export_col2:
    st.caption(
        f"Export akan berisi **{len(detail):,} Unique Code** sesuai filter aktif."
    )

st.dataframe(
    detail,
    use_container_width=True,
    hide_index=True,
)

# -------------------------
# Scope detail
# -------------------------
st.divider()
st.header("Scope Detail")

if len(filtered) == 0:
    st.info("Tidak ada data sesuai filter.")
else:
    selected_code = st.selectbox(
        "Pilih Unique Code",
        filtered["Unique Code"].tolist(),
    )

    selected = filtered[filtered["Unique Code"] == selected_code].iloc[0]

    info1, info2, info3 = st.columns(3)
    info1.write(f"**Contract**  \n{selected.get('Contract', '')}")
    info2.write(f"**Scope**  \n{selected.get('Scope', '')}")
    info3.write(f"**COD**  \n{selected.get('COD', '')}")

    st.subheader("Current Monitoring")
    m1, m2, m3 = st.columns(3)
    m1.metric("Status Pekerjaan", str(selected.get("STATUS PEKERJAAN", "") or "Data Tidak Tersedia"))
    m2.metric("Bidang", str(selected.get("BIDANG", "") or "Data Tidak Tersedia"))
    m3.metric("Lingkup", str(selected.get("LINGKUP", "") or "Data Tidak Tersedia"))

    st.subheader("Source Data Detail")
    source_detail = selected.to_frame("Value")
    source_detail.index.name = "Attribute"
    st.dataframe(source_detail, use_container_width=True)
