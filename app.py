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

if uploaded_files:
    all_data = []

    column_names = [
        "Record_Type",
        "Line_Number",
        "Type_1",
        "Customer_Code",
        "Type_2",
        "Destination_Code",
        "Type_3",
        "Country_Code",
        "Date",
        "Quantity",
        "Reference_Number",
        "Blank_Field",
        "Company_Name",
        "Extra_1",
        "Extra_2",
        "Extra_3",
        "Extra_4",
        "Extra_5",
        "Extra_6",
        "Extra_7",
        "Extra_8",
        "Extra_9",
        "End_Record",
        "File_Name"
    ]

    for file in uploaded_files:
        try:
            df = pd.read_csv(
                file,
                sep="^",
                header=None,
                dtype=str,
                engine="python"
            )

            df = df.iloc[:, :23]
            df.columns = column_names[:-1]

            df["Company_Name"] = df["Company_Name"].str.strip()

            df.insert(len(df.columns), "File_Name", file.name)

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
