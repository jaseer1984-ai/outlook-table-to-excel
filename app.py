import streamlit as st
import pandas as pd
from io import BytesIO
import re
import easyocr
import numpy as np
from PIL import Image
from streamlit_paste_button import paste_image_button

st.set_page_config(page_title="Screenshot to Payment Excel", layout="wide")

st.title("Screenshot to Payment Excel")

st.write("Take screenshot using Win + Shift + S, then click below button to paste.")

OUTPUT_COLUMNS = [
    "BRANCH",
    "BENEFICIARY NAME",
    "BENEFICIARY ACCOUNT NUMBER",
    "SWIFT CODE",
    "DESCRIPTION"
]

@st.cache_resource
def load_reader():
    return easyocr.Reader(["en"], gpu=False)

def clean_text(x):
    return re.sub(r"\s+", " ", str(x)).strip()

def find_value(lines, possible_headers):
    for i, line in enumerate(lines):
        key = re.sub(r"[^a-z0-9]", "", line.lower())
        for h in possible_headers:
            if h in key:
                if i + 1 < len(lines):
                    return clean_text(lines[i + 1])
    return ""

def extract_data_from_ocr(text):
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]

    emp_id = ""
    emp_name = ""
    swift = ""
    iban = ""
    desc = ""
    branch = ""

    for line in lines:
        if re.search(r"\bUNI\d+\b", line.upper()):
            emp_id = re.search(r"\bUNI\d+\b", line.upper()).group(0)

        iban_match = re.search(r"\bSA\d{22}\b", line.replace(" ", "").upper())
        if iban_match:
            iban = iban_match.group(0)

        swift_match = re.search(r"\b[A-Z]{4}SA[A-Z0-9]{2,5}\b", line.replace(" ", "").upper())
        if swift_match:
            swift = swift_match.group(0)

    emp_name = find_value(lines, [
        "employeename",
        "employeenameenglish",
        "beneficiaryname",
        "beneficaryname"
    ])

    desc = find_value(lines, [
        "description",
        "remarks",
        "remark"
    ])

    branch = find_value(lines, [
        "branch",
        "location"
    ])

    if emp_name:
        emp_name = emp_name[:35]

    description_final = (desc + " " + emp_id).strip()

    return pd.DataFrame([{
        "BRANCH": branch,
        "BENEFICIARY NAME": emp_name,
        "BENEFICIARY ACCOUNT NUMBER": iban,
        "SWIFT CODE": swift,
        "DESCRIPTION": description_final
    }], columns=OUTPUT_COLUMNS)

paste_result = paste_image_button(
    label="Paste Screenshot",
    background_color="#1f77b4",
    hover_background_color="#155a8a",
    errors="ignore"
)

if paste_result.image_data is not None:
    st.image(paste_result.image_data, caption="Pasted Screenshot", use_container_width=True)

    with st.spinner("Reading screenshot..."):
        reader = load_reader()
        img = Image.open(BytesIO(paste_result.image_data)) if isinstance(paste_result.image_data, bytes) else paste_result.image_data
        img_np = np.array(img)
        ocr_result = reader.readtext(img_np, detail=0, paragraph=False)
        extracted_text = "\n".join(ocr_result)

    with st.expander("OCR extracted text"):
        st.text(extracted_text)

    final_df = extract_data_from_ocr(extracted_text)

    st.success("Data extracted. Please verify before download.")
    edited_df = st.data_editor(final_df, use_container_width=True, num_rows="dynamic")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        edited_df.to_excel(writer, index=False, sheet_name="Payment")

    st.download_button(
        label="Download Excel",
        data=output.getvalue(),
        file_name="payment_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
