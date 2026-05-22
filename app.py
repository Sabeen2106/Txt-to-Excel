# app.py
import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Text Files to Excel Converter")

uploaded_files = st.file_uploader(
    "Upload text files",
    type=["txt", "csv"],
    accept_multiple_files=True
)

separator = st.selectbox(
    "Choose file separator",
    ["Auto detect", "Comma (,)", "Tab", "Pipe (|)", "Semicolon (;)"]
)

def get_sep(choice):
    return {
        "Comma (,)": ",",
        "Tab": "\t",
        "Pipe (|)": "|",
        "Semicolon (;)": ";"
    }.get(choice, None)

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        try:
            sep = get_sep(separator)

            if sep:
                df = pd.read_csv(file, sep=sep)
            else:
                df = pd.read_csv(file, sep=None, engine="python")

            df["source_file"] = file.name
            all_data.append(df)

        except Exception as e:
            st.error(f"Could not read {file.name}: {e}")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)

        st.success(f"Successfully processed {len(all_data)} files")
        st.dataframe(final_df)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Combined_Data")

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name="combined_text_files.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
