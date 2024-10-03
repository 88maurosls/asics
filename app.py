import streamlit as st
import pandas as pd
import io
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime


# Funzione per formattare la colonna Taglia
def format_taglia(size_us):
    size_str = str(size_us)
    if ".0" in size_str:
        size_str = size_str.replace(".0", "")  # Rimuovi il .0
    return size_str.replace(".5", "+")  # Converti .5 in +

# Funzione per pulire i prezzi (rimuovi simbolo dell'euro e converti in float)
def clean_price(price):
    return float(str(price).replace("€", "").replace(",", "").strip())

# Funzione per duplicare le righe in base al valore di Qta
def expand_rows(df):
    expanded_df = df.loc[df.index.repeat(df['Qta'])].assign(Qta=1)
    expanded_df['Tot Costo'] = expanded_df['Costo']
    return expanded_df

# Funzione per caricare il file color.txt e restituire un dizionario di mapping
def load_colors_mapping(file_path):
    colors_mapping = {}
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if ';' in line:
                try:
                    key, value = line.split(';')
                    colors_mapping[key] = value
                except ValueError:
                    st.warning(f"Errore nel parsing della riga: {line}. Ignorata.")
            else:
                st.warning(f"Riga malformata nel file color.txt: {line}. Ignorata.")
    return colors_mapping

# Funzione per determinare il valore di "Base Color"
def get_base_color(color_name, colors_mapping):
    for key in colors_mapping:
        if color_name.upper().startswith(key):
            return colors_mapping[key]
    return ""  # Se non trovi corrispondenza, lascia vuoto

# Funzione per elaborare ogni file caricato
def process_file(file, colors_mapping, ricarico):
    df = pd.read_excel(file, dtype={'Color code': str, 'EAN code': str})
    
    output_df = pd.DataFrame({
        "Articolo": df["Trading code"],
        "Descrizione": df["Item name"],
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color code"].apply(lambda x: x.zfill(3)),
        "Base Color": df["Color name"].apply(lambda x: get_base_color(x, colors_mapping)),
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": df["Unit price"].apply(clean_price),
        "Retail": df["Unit price"].apply(clean_price) * ricarico,
        "Taglia": df["Size US"].apply(format_taglia),
        "Barcode": df["EAN code"],
        "EAN": "",
        "Qta": df["Quantity"],
        "Tot Costo": df["Unit price"].apply(clean_price) * df["Quantity"] * ricarico,
        "Materiale": "",
        "Spec. Materiale": "",
        "Misure": "",
        "Scala Taglie": "US",
        "Tacco": "",
        "Suola": "",
        "Carryover": "",
        "HS Code": ""
    })
    
    expanded_df = expand_rows(output_df)
    
    return expanded_df

# Funzione per suddividere i dati in fogli di massimo 50 righe e aggiungere l'intestazione
def write_data_in_chunks(writer, df, stagione, data_inizio, data_fine, ricarico):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)
    for i in range(num_chunks):
        chunk_df = df[i*50:(i+1)*50]
        sheet_name = f"Foglio{i+1}"
        start_row = 9
        chunk_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)

        if sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
        else:
            raise ValueError(f"Il foglio {sheet_name} non è stato trovato!")

        worksheet.write('A1', 'STAGIONE:')
        worksheet.write('B1', stagione)
        worksheet.write('A2', 'TIPO:')
        worksheet.write('B2', 'ACCESSORI')
        worksheet.write('A3', 'DATA INIZIO:')
        worksheet.write('B3', data_inizio.strftime('%d/%m/%Y'))
        worksheet.write('A4', 'DATA FINE:')
        worksheet.write('B4', data_fine.strftime('%d/%m/%Y'))
        worksheet.write('A5', 'RICARICO:')
        worksheet.write('B5', ricarico)

        text_format = writer.book.add_format({'num_format': '@'})
        worksheet.set_column('L:L', 20, text_format)

        last_data_row = len(chunk_df) + start_row
        empty_row = last_data_row + 1
        worksheet.write(f'N{empty_row}', "")
        worksheet.write(f'O{empty_row}', "")

        total_row = empty_row + 2
        number_format = writer.book.add_format({'num_format': '#,##0.00'})
        worksheet.write_formula(f'N{total_row}', f"=SUM(N{start_row+2}:N{last_data_row + 1})", number_format)
        worksheet.write_formula(f'O{total_row}', f"=SUM(O{start_row+2}:O{last_data_row + 1})", number_format)
        worksheet.set_column('N:N', None, number_format)
        worksheet.set_column('O:O', None, number_format)

# Funzione per connettersi a Google Sheets
def connect_to_gsheet():
    # Usa i segreti in formato TOML salvati su Streamlit
    credentials = {
        "type": st.secrets["gcp_service_account"]["type"],
        "project_id": st.secrets["gcp_service_account"]["project_id"],
        "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
        "private_key": st.secrets["gcp_service_account"]["private_key"],
        "client_email": st.secrets["gcp_service_account"]["client_email"],
        "client_id": st.secrets["gcp_service_account"]["client_id"],
        "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
        "token_uri": st.secrets["gcp_service_account"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"]
    }

    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(credentials, scopes=scope)
    client = gspread.authorize(creds)
    return client

# Funzione per scrivere dati su Google Sheets
def write_to_gsheet(data, sheet_url):
    client = connect_to_gsheet()
    # Apri il Google Sheet con l'URL fornito
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)  # Seleziona il primo foglio
    
    # Trova la prima riga vuota
    next_row = len(worksheet.get_all_values()) + 1
    
    # Scrivi i dati nella colonna specificata
    for index, (articolo, colore, gender) in enumerate(data):
        worksheet.update(f'A{next_row + index}', articolo)  # Colonna per Articolo
        worksheet.update(f'B{next_row + index}', colore)    # Colonna per Colore
        worksheet.update(f'C{next_row + index}', gender)    # Colonna per Gender

# Streamlit app
st.title('Asics Xmag Lineare')

# Campi di input per l'intestazione
stagione = st.text_input("Inserisci STAGIONE")
data_inizio = st.date_input("Inserisci DATA INIZIO")
data_fine = st.date_input("Inserisci DATA FINE")
ricarico = st.text_input("Inserisci RICARICO", value="2")  # Imposta 2 come valore predefinito

# Aggiungi il contenuto testuale con il link
st.markdown('**[Scarica le Packing List da qui](https://b2b.asics.com/orders-overview/order-history)**')

# Carica il file color.txt dalla directory del progetto
colors_mapping = load_colors_mapping("color.txt")

# Permetti l'upload di più file Excel
uploaded_files = st.file_uploader("Scegli i file Excel", accept_multiple_files=True)

if uploaded_files and stagione and data_inizio and data_fine and ricarico:
    ricarico = float(ricarico)  # Converte RICARICO in float
    processed_dfs = []
    
    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file, colors_mapping, ricarico))
    
    final_df = pd.concat(processed_dfs, ignore_index=True)

    unique_combinations = final_df[["Articolo", "Colore"]].drop_duplicates()

    st.write("Anteprima Articolo-Colore:")

    selections = {}

    for index, row in unique_combinations.iterrows():
        # Aggiungi l'opzione "UNISEX" al selectbox
        flag = st.selectbox(f"{row['Articolo']}-{row['Colore']}", options=["Seleziona...", "UOMO", "DONNA", "UNISEX"], key=index)
        selections[(row['Articolo'], row['Colore'])] = flag

    if st.button("Elabora File"):
        if any(flag == "Seleziona..." for flag in selections.values()):
            st.error("Devi selezionare UOMO, DONNA o UNISEX per tutte le combinazioni!")
        else:
            # Prepara i dati da inviare a Google Sheets
            gsheet_data = [(row['Articolo'], row['Colore'], selections[(row['Articolo'], row['Colore'])]) for index, row in unique_combinations.iterrows()]
            
            # Scrivi i dati nel Google Sheet
            google_sheet_url = "https://docs.google.com/spreadsheets/d/1p84nF9Tq-1ZJgQSEJcgrePLvQyGQ3cjt_1IZP5qPs00/edit?usp=sharing"
            write_to_gsheet(gsheet_data, google_sheet_url)

            st.success("Dati scritti con successo su Google Sheet!")
