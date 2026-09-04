import re
import io

import pandas as pd
import streamlit as st
from pypdf import PdfReader


HOURS_PER_MD = 8


def extract_text_from_pdf(uploaded_file):
    reader = PdfReader(uploaded_file)
    text = ""

    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"

    return text


def parse_invoice(text):
    text_clean = text.replace(",", "")

    # Mohammad Haidar
    if "Mohammad Haidar" in text or "MOHAMAD HUSSEIN HAIDAR" in text:
        name = "Mohammad Haidar"

        days_match = re.search(
            r"Charged\s+([\d.]+)\s+days",
            text,
            re.IGNORECASE
        )

        total_match = re.search(
            r"TOTAL VALUE\s*([\d,]+)",
            text,
            re.IGNORECASE
        )

        mds = float(days_match.group(1)) if days_match else None

        if total_match:
            total = float(total_match.group(1).replace(",", ""))
        else:
            total = None

        return {
            "Jméno": name,
            "Počet MDs": mds,
            "Celková částka": total,
        }

    # Santhosh Balakrishnan
    if "Santhosh Balakrishnan" in text:
        name = "Santhosh Balakrishnan"

        match = re.search(
            r"Software Development\s+([\d.]+).*?([\d,]+\.\d{2})",
            text,
            re.IGNORECASE | re.DOTALL,
        )

        if match:
            mds = float(match.group(1))
            total = float(match.group(2).replace(",", ""))
        else:
            mds = None
            total = None

        return {
            "Jméno": name,
            "Počet MDs": mds,
            "Celková částka": total,
        }

    # Laman Aghabayova
    if "AGHABAYOVA" in text.upper():
        name = "Laman Aghabayova"

        hours_match = re.search(
            r"([\d,.]+)\s*hours",
            text,
            re.IGNORECASE,
        )

        total_match = re.search(
            r"Total to pay\s*\(EUR\)\s*([\d\s,.]+)",
            text,
            re.IGNORECASE,
        )

        if hours_match:
            hours_text = hours_match.group(1)
            hours = float(
                hours_text.replace(" ", "").replace(",", ".")
            )
            mds = hours / HOURS_PER_MD
        else:
            mds = None

        if total_match:
            total_text = total_match.group(1)
            total_text = (
                total_text
                .replace(" ", "")
                .replace(",", ".")
            )
            total = float(total_text)
        else:
            total = None

        return {
            "Jméno": name,
            "Počet MDs": mds,
            "Celková částka": total,
        }

    return {
        "Jméno": "Nerozpoznáno",
        "Počet MDs": None,
        "Celková částka": None,
    }


st.set_page_config(
    page_title="Invoice Automation",
    page_icon="📄",
)

st.title("Invoice Automation")

st.write(
    "Nahraj faktury a aplikace z nich vytvoří přehled pro Excel."
)

uploaded_files = st.file_uploader(
    "Přetáhni sem PDF faktury",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    results = []

    for uploaded_file in uploaded_files:
        text = extract_text_from_pdf(uploaded_file)
        result = parse_invoice(text)
        results.append(result)

    df = pd.DataFrame(results)

    st.subheader("Výsledek")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Invoices",
        )

    output.seek(0)

    st.download_button(
        label="Stáhnout Excel",
        data=output,
        file_name="invoice_output.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )