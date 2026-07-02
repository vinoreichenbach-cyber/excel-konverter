from __future__ import annotations

from datetime import datetime
from pathlib import Path
import tempfile

import streamlit as st

from converter import (
    create_output_workbook,
    extract_source_table,
    infer_customer_info,
    build_output_rows,
)

st.set_page_config(page_title="Excel-Konverter", page_icon="📊", layout="centered")

st.title("Excel-Konverter")
st.write("AMS-Export hochladen und direkt als formatierte KFZ-Beitragsübersicht herunterladen.")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xls", "xlsx"])

with st.expander("Optionale Angaben"):
    stand = st.text_input("Stand", value=datetime.now().strftime("%m.%Y"))
    kunde = st.text_input("Versicherungsnehmer", value="")
    adresse = st.text_input("Adresse", value="")
    ort = st.text_input("PLZ und Ort", value="")
    logo_file = st.file_uploader("Logo optional hochladen", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    try:
        with tempfile.TemporaryDirectory(prefix="excel_konverter_") as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            input_path = temp_dir / uploaded_file.name
            output_path = temp_dir / "amsexport_formatiert.xlsx"

            input_path.write_bytes(uploaded_file.getvalue())

            logo_path = None
            if logo_file is not None:
                logo_path = temp_dir / logo_file.name
                logo_path.write_bytes(logo_file.getvalue())

            source_table = extract_source_table(input_path)
            customer_info = infer_customer_info(source_table)
            output_rows = build_output_rows(source_table)

            final_kunde = kunde or customer_info.name
            final_adresse = adresse or customer_info.street
            final_ort = ort or " ".join(
                part for part in (customer_info.zip_code, customer_info.city) if part
            )

            create_output_workbook(
                output_path=output_path,
                output_rows=output_rows,
                stand=stand,
                kunde=final_kunde,
                adresse=final_adresse,
                ort=final_ort,
                logo_path=logo_path,
            )

            st.success(f"Fertig. {len(output_rows)} Zeilen wurden verarbeitet.")

            with output_path.open("rb") as file:
                st.download_button(
                    label="Formatierte Excel herunterladen",
                    data=file.read(),
                    file_name="amsexport_formatiert.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except Exception as exc:
        st.error("Die Datei konnte nicht umgewandelt werden.")
        st.exception(exc)
