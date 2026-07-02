import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel-Konverter")

datei = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])

if datei:
    df = pd.read_excel(datei)

    st.write("Vorschau:")
    st.dataframe(df)

    # Hier kommt später deine Umwandlung rein
    df["Bearbeitet"] = "Ja"

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        "Umgewandelte Excel herunterladen",
        output.getvalue(),
        file_name="umgewandelt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
