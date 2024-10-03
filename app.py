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
    credentials = {
        "type": st.secrets["gsheet"]["type"],
        "project_id": st.secrets["gsheet"]["project_id"],
        "private_key_id": st.secrets["gsheet"]["private_key_id"],
        "private_key": st.secrets["gsheet"]["private_key"],
        "client_email": st.secrets["gsheet"]["client_email"],
        "client_id": st.secrets["gsheet"]["client_id"],
        "auth_uri": st.secrets["gsheet"]["auth_uri"],
        "token_uri": st.secrets["gsheet"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gsheet"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gsheet"]["client_x509_cert_url"]
    }

    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(credentials, scopes=scope)
    client = gspread.authorize(creds)
    return client

# Funzione per recuperare i dati "Articolo", "Colore" e "Gender" dal foglio "Gender"
def get_existing_gender(sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Gender")
    
    # Recupera tutti i valori dal foglio
    data = worksheet.get_all_values()

    # Creare un dizionario {"Articolo-Colore": Gender} per compatibilità JSON
    gender_dict = {f"{row[0]}-{row[1]}": row[2] for row in data[1:]}  # Ignora l'intestazione
    return gender_dict

# Funzione per scrivere o aggiornare dati su Google Sheets
def write_to_gsheet(data, sheet_url):
    client = connect_to_gsheet()
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.worksheet("Gender")

    # Recupera i dati esistenti dal foglio
    existing_data = worksheet.get_all_values()

    # Ottieni le combinazioni già esistenti (Articolo, Colore) con il loro indice di riga
    existing_entries = {f"{row[0]}-{row[1]}": idx+2 for idx, row in enumerate(existing_data[1:])}  # Ignora l'intestazione

    # Prepara gli aggiornamenti in batch
    batch_updates = []

    for (articolo, colore, gender) in data:
        key = f"{articolo}-{colore}"

        if key in existing_entries:
            # Se la combinazione esiste già, aggiorna la riga esistente
            row_to_update = existing_entries[key]
            batch_updates.append({
                "range": f'C{row_to_update}',  # Aggiorna solo la colonna C (Gender)
                "values": [[gender]]
            })
        else:
            # Altrimenti, aggiungi una nuova riga
            worksheet.append_row([articolo, colore, gender])

    if batch_updates:
        # Esegui l'aggiornamento in batch
        worksheet.batch_update(batch_updates)

    st.success(f"Dati aggiornati o aggiunti su Google Sheet.")

# Streamlit app e scrittura del file
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
    
    # Recupera il genere già presente nel foglio "Gender"
    google_sheet_url = "https://docs.google.com/spreadsheets/d/1p84nF9Tq-1ZJgQSEJcgrePLvQyGQ3cjt_1IZP5qPs00/edit?usp=sharing"
    gender_dict = get_existing_gender(google_sheet_url)
    
    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file, colors_mapping, ricarico))
    
    final_df = pd.concat(processed_dfs, ignore_index=True)

    unique_combinations = final_df[["Articolo", "Colore"]].drop_duplicates()

    st.write("Anteprima Articolo-Colore:")

    selections = {}

    for index, row in unique_combinations.iterrows():
        articolo_colore = f"{row['Articolo']}-{row['Colore']}"  # Chiave come stringa unica
        
        # Se il genere è già presente in Google Sheet, usalo, altrimenti usa "Seleziona..."
        preselected_gender = gender_dict.get(articolo_colore, "Seleziona...")

        flag = st.selectbox(
            f"{row['Articolo']}-{row['Colore']}", 
            options=["Seleziona...", "UOMO", "DONNA", "UNISEX"], 
            key=index, 
            index=["Seleziona...", "UOMO", "DONNA", "UNISEX"].index(preselected_gender) 
            if preselected_gender in ["UOMO", "DONNA", "UNISEX"] else 0
        )
        selections[(row['Articolo'], row['Colore'])] = flag

    if st.button("Elabora File"):
        if any(flag == "Seleziona..." for flag in selections.values()):
            st.error("Devi selezionare UOMO, DONNA o UNISEX per tutte le combinazioni!")
        else:
            # Prepara i dati da inviare a Google Sheets
            gsheet_data = [(row['Articolo'], row['Colore'], selections[(row['Articolo'], row['Colore'])]) for index, row in unique_combinations.iterrows()]
            
            # Scrivi i dati nel Google Sheet
            write_to_gsheet(gsheet_data, google_sheet_url)

            # Dividi i dati per genere
            uomo_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UOMO', axis=1)]
            donna_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'DONNA', axis=1)]
            unisex_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UNISEX', axis=1)]

            uomo_output = io.BytesIO()
            donna_output = io.BytesIO()
            unisex_output = io.BytesIO()

            # Genera file Excel per UOMO
            if not uomo_df.empty:
                with pd.ExcelWriter(uomo_output, engine='xlsxwriter') as writer_uomo:
                    write_data_in_chunks(writer_uomo, uomo_df, stagione, data_inizio, data_fine, ricarico)
                st.download_button(
                    label="Download File UOMO",
                    data=uomo_output.getvalue(),
                    file_name="uomo_processed_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Genera file Excel per DONNA
            if not donna_df.empty:
                with pd.ExcelWriter(donna_output, engine='xlsxwriter') as writer_donna:
                    write_data_in_chunks(writer_donna, donna_df, stagione, data_inizio, data_fine, ricarico)
                st.download_button(
                    label="Download File DONNA",
                    data=donna_output.getvalue(),
                    file_name="donna_processed_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Genera file Excel per UNISEX
            if not unisex_df.empty:
                with pd.ExcelWriter(unisex_output, engine='xlsxwriter') as writer_unisex:
                    write_data_in_chunks(writer_unisex, unisex_df, stagione, data_inizio, data_fine, ricarico)
                st.download_button(
                    label="Download File UNISEX",
                    data=unisex_output.getvalue(),
                    file_name="unisex_processed_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
