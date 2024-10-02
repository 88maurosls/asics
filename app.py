import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Funzione per autenticarsi e aprire il Google Sheet con chiave e email
def connect_to_google_sheet():
    try:
        # Imposta lo scope per le API di Google Sheets e Google Drive
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

        # Utilizza la chiave e l'email che hai fornito
        creds = {
            "type": "service_account",
            "project_id": "innate-entry-434509",
            "private_key_id": "591e247304cab81eb4feebec07d0903bd5f05fcc",  # La tua chiave privata
            "private_key": """-----BEGIN PRIVATE KEY-----\n<INSERISCI QUI LA CHIAVE>\n-----END PRIVATE KEY-----\n""",  # Sostituisci con la tua chiave privata
            "client_email": "marcatempo@innate-entry-434509-r0.iam.gserviceaccount.com",
            "client_id": "<INSERISCI QUI IL CLIENT ID>",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/marcatempo%40innate-entry-434509-r0.iam.gserviceaccount.com"
        }

        # Usa la chiave e l'email per autorizzare l'accesso al Google Sheet
        client = gspread.service_account_from_dict(creds, scopes=scope)
        # Apri il Google Sheet tramite il link
        sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/17g287VcrlbQfO9u6cmySRR8sAzzMUV2hjWe_aSYr3x8/edit?usp=sharing')
        return sheet
    except Exception as e:
        st.error(f"Errore durante la connessione al Google Sheet: {e}")
        return None

# Funzione per scrivere i dati di genere nel Google Sheet
def update_google_sheet(gender_dict):
    try:
        sheet = connect_to_google_sheet()
        if sheet:
            worksheet = sheet.get_worksheet(0)  # Primo foglio nel documento

            # Aggiorna il foglio con i dati di genere
            for (articolo, colore), flag in gender_dict.items():
                worksheet.append_row([articolo, colore, flag])  # Aggiungi una nuova riga con i dati
            st.success("Google Sheet aggiornato correttamente.")
        else:
            st.error("Impossibile connettersi al Google Sheet.")
    except Exception as e:
        st.error(f"Errore durante l'aggiornamento del Google Sheet: {e}")

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

    # Seleziona solo le combinazioni uniche di "Articolo" e "Colore"
    unique_combinations = final_df[["Articolo", "Colore"]].drop_duplicates()

    st.write("Anteprima delle combinazioni uniche di Articolo e Colore:")

    # Dizionario per raccogliere il flag UOMO/DONNA per ogni combinazione
    selections = {}

    # Visualizzare l'anteprima dei dati unici con opzione per UOMO/DONNA
    for index, row in unique_combinations.iterrows():
        flag = st.radio(f"Articolo: {row['Articolo']} - Colore: {row['Colore']}", ('UOMO', 'DONNA'), key=index)
        selections[(row['Articolo'], row['Colore'])] = flag

    if st.button("Elabora File"):
        # Filtra i dati in base alla selezione UOMO/DONNA
        uomo_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UOMO', axis=1)]
        donna_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'DONNA', axis=1)]

        # Aggiorna il Google Sheet con i dati di genere
        update_google_sheet(selections)

        # Crea un file in memoria per il download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Scrivi i dati in due fogli separati
            uomo_df.to_excel(writer, sheet_name='UOMO', index=False)
            donna_df.to_excel(writer, sheet_name='DONNA', index=False)
            
            # Imposta la colonna "Barcode" come testo per evitare la notazione scientifica
            worksheet_uomo = writer.sheets['UOMO']
            worksheet_donna = writer.sheets['DONNA']
            text_format = writer.book.add_format({'num_format': '@'})  # Formato per trattare come testo
            worksheet_uomo.set_column('L:L', 20, text_format)  # Formatta la colonna Barcode come testo
            worksheet_donna.set_column('L:L', 20, text_format)  # Formatta la colonna Barcode come testo

        # Fornisci un pulsante per scaricare il file elaborato
        st.download_button(
            label="Download Processed Excel",
            data=output.getvalue(),
            file_name="processed_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
