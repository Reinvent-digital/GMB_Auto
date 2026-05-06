import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

from gmb_utils import (
    METRIC_COLUMNS,
    METRIC_DESCRIPTIONS,
    build_comparison_table,
    infer_month_label,
    normalize_columns,
    save_comparison_excel
)

st.set_page_config(
    page_title="GMB Insights Comparator",
    page_icon="📈",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# --- Custom CSS for a better look ---
st.markdown("""
<style>
    .main {
        background-color: #f9fbfd;
    }
    h1 {
        color: #1f3b73;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 800;
    }
    div[data-testid="stFileUploader"] {
        background-color: white;
        padding: 1.5rem;
        border-radius: 0.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    .stDownloadButton > button {
        background-color: #4CAF50;
        color: white;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stDownloadButton > button:hover {
        background-color: #45a049;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }
</style>
""", unsafe_allow_html=True)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Reads a Streamlit UploadedFile into a clean GMB DataFrame."""
    suffix = Path(uploaded_file.name).suffix.lower()
    
    # Needs to reset pointer if read multiple times
    uploaded_file.seek(0)
    
    if suffix in (".xlsx", ".xls"):
        raw = pd.read_excel(uploaded_file, header=[0, 1], dtype=str)
    else:
        raw = pd.read_csv(uploaded_file, header=[0, 1], dtype=str)

    cols = normalize_columns(raw.columns)
    raw.columns = cols

    if "Business name" not in raw.columns:
        uploaded_file.seek(0)
        if suffix in (".xlsx", ".xls"):
            raw = pd.read_excel(uploaded_file, header=0, dtype=str)
        else:
            raw = pd.read_csv(uploaded_file, header=0, dtype=str)
        raw.columns = normalize_columns(raw.columns)

    keep = [c for c in raw.columns if c in ("Business name", "Address") + tuple(METRIC_COLUMNS)]
    if "Business name" not in keep:
        raise ValueError(f"Could not find 'Business name' column in {uploaded_file.name}")

    df = raw[keep].copy()
    df["Business name"] = df["Business name"].astype(str).str.strip()
    
    for metric in METRIC_COLUMNS:
        if metric in df.columns:
            df[metric] = pd.to_numeric(df[metric].fillna(0).replace("", 0), errors="coerce").fillna(0).astype(int)
            
    return df


def make_display_dataframe(df: pd.DataFrame, labels: list[str]) -> pd.DataFrame:
    """Formats the flat dataframe into a MultiIndex dataframe for a beautiful Streamlit display"""
    tuples = []
    if "Business name" in df.columns:
        tuples.append(("Business name", "", ""))
    if "Address" in df.columns:
        tuples.append(("Address", "", ""))
        
    for col in df.columns:
        if col in ("Business name", "Address"):
            continue
        # Extract metric and label
        for label in labels:
            ext = f" {label}"
            if col.endswith(ext):
                metric = col[:-len(ext)]
                desc = METRIC_DESCRIPTIONS.get(metric, "")
                tuples.append((metric, desc, label))
                break
    
    df_mi = df.copy()
    df_mi.columns = pd.MultiIndex.from_tuples(tuples)
    return df_mi


def main():
    st.title("📈 GMB Insights Comparator")
    st.markdown("Upload your monthly Google Business Profile reports to automatically generate a side-by-side comparison.")

    st.markdown("---")

    # File uploads
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("📁 Month 1")
        file1 = st.file_uploader("Upload first report", type=["csv", "xlsx", "xls"], key="file1")
        label1 = st.text_input("Label (e.g. February)", placeholder="Auto-detect", key="label1")

    with col2:
        st.subheader("📁 Month 2")
        file2 = st.file_uploader("Upload second report", type=["csv", "xlsx", "xls"], key="file2")
        label2 = st.text_input("Label (e.g. March)", placeholder="Auto-detect", key="label2")
        
    with col3:
        st.subheader("📁 Month 3 (Optional)")
        file3 = st.file_uploader("Upload third report", type=["csv", "xlsx", "xls"], key="file3")
        label3 = st.text_input("Label (e.g. April)", placeholder="Auto-detect", key="label3")

    st.markdown("---")
    
    # Options
    col_opt1, col_opt2 = st.columns([1, 2])
    with col_opt1:
        output_format = st.radio("Choose Output Format:", ["Excel (.xlsx) - Styled", "CSV - Raw Data"])

    # Processing
    if file1 and file2:
        try:
            with st.spinner("Processing reports & merging data..."):
                datasets = []
                labels = []
                
                # Helper to process uploads
                def add_dataset(file_obj, raw_label):
                    final_label = raw_label.strip() if raw_label.strip() else infer_month_label(Path(file_obj.name))
                    df = read_uploaded_file(file_obj)
                    datasets.append((df, final_label))
                    labels.append(final_label)

                add_dataset(file1, label1)
                add_dataset(file2, label2)
                if file3:
                    add_dataset(file3, label3)

                # Process
                result = build_comparison_table(datasets)
                
                st.success(f"Successfully matched **{len(result)}** locations across {len(datasets)} months!")
                
                # Show a mini preview
                st.write("### Data Preview")
                display_df = make_display_dataframe(result, labels)
                st.dataframe(display_df.head(), width="stretch")

                buffer = BytesIO()
                
                if output_format.startswith("Excel"):
                    save_comparison_excel(result, buffer, labels)
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    file_ext = "xlsx"
                else:
                    result.to_csv(buffer, index=False)
                    mime_type = "text/csv"
                    file_ext = "csv"
                    
                buffer.seek(0)
                file_name = f"gmb_comparison_{'_vs_'.join(labels)}.{file_ext}".replace(" ", "_")

                
                # Big Download Button
                st.download_button(
                    label=f"⬇️ Download {file_ext.upper()} Report",
                    data=buffer,
                    file_name=file_name,
                    mime=mime_type,
                    width="stretch"
                )
                
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            
    else:
        st.info("👆 Please upload both files to generate your comparison.")


if __name__ == "__main__":
    main()
