#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, sys, traceback, json
from pathlib import Path

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

os.chdir(HERE)

from generate import run, parse_ycsx

st.set_page_config(page_title="Auto Định Mức — Test Tool", layout="centered")
st.title("🧪 Auto Định Mức — Test")

st.markdown(
    "Upload an order file (.json or .xlsx YCSX) and enter the number of print colors. "
    "The tool will generate the Định mức + YCSX files. "
    "Any errors are shown in full — this is a **test surface**."
)

uploaded_file = st.file_uploader(
    "Order file", type=["json", "xlsx"],
    help="Upload a JSON order or a YCSX .xlsx production request form"
)

colors = st.number_input("Số màu in (number of print colors)", min_value=0, max_value=12, value=3)

if st.button("Generate", type="primary"):
    if not uploaded_file:
        st.error("Please upload an order file first.")
        st.stop()

    with st.spinner("Generating..."):
        try:
            suffix = Path(uploaded_file.name).suffix.lower()
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = os.path.join(tmp, uploaded_file.name)
                with open(tmp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                if suffix == ".json":
                    source = ("order", tmp_path)
                elif suffix == ".xlsx":
                    source = ("ycsx", tmp_path)
                else:
                    st.error(f"Unsupported file type: {suffix}")
                    st.stop()

                res = run(source, colors, outdir=os.path.join(tmp, "out"))

            st.success(f"Bag family: **{res['family']}** | Products: **{len(res['products'])}**")
            for p in res["products"]:
                f = p["fields"]
                with st.expander(f"📦 {p['product_name']}"):
                    for k in ("qty", "bag_length_m", "width_plus_gusset_m",
                              "sl_in_thuc_te_m", "so_mau_in",
                              "kho_manh", "kho_mang", "kho_giay",
                              "inner_bag_weight_kg", "mang_bopp_kg",
                              "dung_moai_opp_kg", "dung_moai_ea_kg",
                              "giay_kraft_kg", "glue_total_kg"):
                        v = f.get(k)
                        if v is not None:
                            st.code(f"{k:24s} = {v}")

            st.subheader("📥 Download outputs")
            for path in res["outputs"]:
                name = os.path.basename(path)
                with open(path, "rb") as f:
                    st.download_button(f"⬇ {name}", data=f, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception:
            st.error("### ❌ Error occurred")
            st.code(traceback.format_exc(), language="python")

st.markdown("---")
st.markdown(
    "**How to deploy on Streamlit Cloud:** "
    "Push this file to GitHub → go to https://streamlit.io/cloud → "
    "connect your repo → set entry point to `test_app.py` → done."
)
