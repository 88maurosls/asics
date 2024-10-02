import streamlit as st
import pandas as pd
import io

# Funzione per formattare la colonna Taglia
def format_taglia(size_us):
    size_str = str(size_us)
    if ".0" in size_str:
        size_str = size_str.replace(".0", "")  # Rimuovi il .0
    return size_str.replace(".5", "+")  # Converti .5 in +

# Funzione per pulire i prezzi (rimuovi simbolo dell'euro e converti in float)
def clean_price(price):
    return float(str(price).replace("€", "").replace(",", "").strip())

# Funzione per elaborare ogni file caricato
def process_file(file):
    df = pd.read_excel(file, dtype={'Color code': str, 'EAN code': str})  # Leggi "Color code" e "EAN code" come stringhe
    
    # Crea il DataFrame di output con le colonne richieste
    output_df = pd.DataFrame({
        "Articolo": df["Trading code"],
        "Descrizione": df["Item name"],
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color code"].apply(lambda x: x.zfill(3)),  # Mantieni gli zeri iniziali
        "Base Color": "",
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": df["Unit price"].apply(clean_price),  # Pulisci e converti i prezzi in numeri
        "Retail": df["Unit price"].apply(clean_price) * 2,  # Moltiplica per 2
        "Taglia": df["Size US"].apply(format_taglia),  # Formatta la colonna Taglia
        "Barcode": df["EAN code"],  # Tratta il barcode come stringa
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

# Streamlit app e scrittura del file
st.title('Upload and Process Multiple Files')

# Permetti l'upload di più file
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True)

if uploaded_files:
    processed_dfs = []
    
    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file))
    
    # Concatenate tutti i DataFrame
    final_df = pd.concat(processed_dfs, ignore_index=True)
    
    # Crea un file in memoria per il download
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
        
        # Imposta la colonna "Barcode" come testo per evitare la notazione scientifica
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column('L:L', 20, None)  # Estendi la larghezza della colonna per il Barcode
        for idx, col in enumerate(final_df):  # Itera sulle colonne per trovare quella del "Barcode"
            if col == "Barcode":
                worksheet.set_column(idx, idx, 20, {'num_format': '@'})  # Formatta come testo la colonna Barcode
    
    # Fornisci un pulsante per scaricare il file elaborato
    st.download_button(
        label="Download Processed Excel",
        data=output.getvalue(),
        file_name="processed_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
