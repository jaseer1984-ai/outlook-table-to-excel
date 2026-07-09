import streamlit as st
import pandas as pd
from io import StringIO, BytesIO

st.set_page_config(page_title="Copy Paste Table to Excel", layout="wide")

st.title("Outlook Table Copy-Paste to Excel")

st.write("Copy the table from Outlook email and paste here. Columns can be different each time.")

paste_text = st.text_area("Paste copied table here", height=300)

def read_pasted_table(text):
    text = text.strip()
    if not text:
        return pd.DataFrame()

    # Try tab-separated first — Outlook/Excel table copy usually works like this
    try:
        df = pd.read_csv(StringIO(text), sep="\t", dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except:
        pass

    # Try comma separated
    try:
        df = pd.read_csv(StringIO(text), dtype=str)
        if df.shape[1] > 1:
            return df.fillna("")
    except:
        pass

    # Fallback: split by multiple spaces
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    rows = [line.split() for line in lines]
    return pd.DataFrame(rows).fillna("")

if st.button("Convert"):
    df = read_pasted_table(paste_text)

    if df.empty:
        st.error("No data found. Please paste the copied table.")
    else:
        st.success("Table detected. You can edit before download.")
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_df.to_excel(writer, index=False, sheet_name="Data")

        st.download_button(
            "Download Excel",
            output.getvalue(),
            "converted_table.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
