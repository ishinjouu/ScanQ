import hashlib
import traceback
import importlib
try:
    import streamlit as st
except ImportError:
    st = None

# Flask API
from flask import Flask, request, jsonify
from utils import sanitize_for_json

app = Flask(__name__)

# ---------------------------------------
# AUTO SWITCH MODULE: DEFAULT / SPECIAL
# ---------------------------------------
def load_processor_module(filename: str):
    filename = filename.lower().strip()
    if filename.startswith("s-"):
        module_name = "special"
    elif filename.startswith("x-"):
        module_name = "InspectionStandard"
    else:
        module_name = "default"
    module = importlib.import_module(module_name)
    return (
        module_name,
        module.extract_table_from_pdf,
        module.bersihkan_dataframe,
        module.transform_to_final_format
    )

# ---------------------------------------
# HASH & UI for streamlit
# ---------------------------------------
def get_file_hash(file):
    return hashlib.md5(file.getvalue()).hexdigest()

st.set_page_config(page_title="Check Sheet QFORM", layout="wide")
st.title("📄 DEBUG CHECK SHEET SCAN QFORM")

uploaded_file = st.file_uploader("📤 Upload PDF file", type="pdf")

if uploaded_file:
    # pilih modul berdasarkan nama file
    module_name, extract_table_from_pdf, bersihkan_dataframe, transform_to_final_format = \
        load_processor_module(uploaded_file.name)
    st.info(f"🧠 Module used: {module_name}")
    
    file_hash = get_file_hash(uploaded_file)

    if "last_file_hash" not in st.session_state or st.session_state.last_file_hash != file_hash:
        st.session_state.last_file_hash = file_hash
        for key in ["df_final_data", "show_updated_table"]:
            st.session_state.pop(key, None)
        st.cache_data.clear()

    try:
        df = extract_table_from_pdf(uploaded_file)
        if df.empty:
            st.warning("❌ No tables detected in the PDF.")
        else:
            df_cleaned = bersihkan_dataframe(df.copy())
            st.subheader("🚀 Cleaned Table")
            st.dataframe(df_cleaned, use_container_width=True, hide_index=True)

            if "df_final_data" not in st.session_state:
                st.session_state.df_final_data = transform_to_final_format(df_cleaned)

            # VALIDASI
            total = len(st.session_state.df_final_data)
            st.info(f"Total rows: {total}")

            st.subheader("📊 Final Format")
            st.dataframe(
                st.session_state.df_final_data.reset_index(drop=True),
                use_container_width=True,
                hide_index=True
            )

            if st.button("♻️ Reset"):
                st.cache_data.clear()
                for key in ["df_final_data", "show_updated_table", "last_file_hash"]:
                    st.session_state.pop(key, None)
                st.rerun()

    except Exception as e:
        st.error(f"🔥 Error:\n{e}")
        st.text(traceback.format_exc())

# ---------------------------------------
# FLASK API
# ---------------------------------------
@app.route("/api/proses_file", methods=["POST"])
def proses_file():
    if 'file' not in request.files:
        return jsonify({"error": "❌ Tidak ada file dikirim"}), 400
    file = request.files['file']

    # pilih modul berdasarkan nama file
    extract_table_from_pdf, bersihkan_dataframe, transform_to_final_format = load_processor_module(file.filename)

    try:
        df_raw = extract_table_from_pdf(file)
        df_clean = bersihkan_dataframe(df_raw)
        df_final = transform_to_final_format(df_clean)
        return jsonify(df_final.to_dict(orient="records"))
    except Exception as e:
        return jsonify({"error": str(e)}), 500
