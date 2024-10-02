import streamlit as st
import pandas as pd
import io

# Function to process each uploaded file
def process_file(file):
    df = pd.read_excel(file, dtype={'Color code': str})  # Leggi la colonna "Color code" come stringa
    
    # Create the output DataFrame with the required columns
    output_df = pd.DataFrame({
        "Articolo": df["Trading code"],
        "Descrizione": df["Item name"],
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color code"].apply(lambda x: x.zfill(3)),  # Mantieni gli zeri iniziali
        "Base Color": "",
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": df["Unit price"],
        "Retail": df["Unit price"] * 2,
        "Taglia": df["Size US"].apply(lambda x: str(x).replace(".5", "+")),
        "Barcode": df["EAN code"],
        "Qta": df["Quantity"],
        "Tot Costo": "",
        "Materiale": "",
        "Spec. Materiale": "",
        "Misure": "",
        "Scala Taglie": "US",
        "Tacco": "",
        "Suola": "",
        "Carryover": "",
        "HS Code": ""
    })
    
    return output_df


# Streamlit app starts here
st.title('Upload and Process Multiple Files')

# Allow the user to upload multiple files
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True)

if uploaded_files:
    # List to store all the processed DataFrames
    processed_dfs = []
    
    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file))
    
    # Concatenate all the processed dataframes
    final_df = pd.concat(processed_dfs, ignore_index=True)
    
    # Create an in-memory file for downloading
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
    
    # Provide a download button
    st.download_button(
        label="Download Processed Excel",
        data=output.getvalue(),
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

