import streamlit as st
import pandas as pd
import io
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
    # Aggiorna la colonna Tot Costo moltiplicando Costo per la nuova Qta (che ora è sempre 1)
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
def process_file(file, colors_mapping):
    df = pd.read_excel(file, dtype={'Color code': str, 'EAN code': str})  # Leggi "Color code" e "EAN code" come stringhe
    
    # Crea il DataFrame di output con le colonne richieste
    output_df = pd.DataFrame({
        "Articolo": df["Trading code"],
        "Descrizione": df["Item name"],
        "Categoria": "CALZATURE",
        "Subcategoria": "Sneakers",
        "Colore": df["Color code"].apply(lambda x: x.zfill(3)),  # Mantieni gli zeri iniziali
        "Base Color": df["Color name"].apply(lambda x: get_base_color(x, colors_mapping)),  # Assegna il "Base Color"
        "Made in": "",
        "Sigla Bimbo": "",
        "Costo": df["Unit price"].apply(clean_price),  # Pulisci e converti i prezzi in numeri
        "Retail": df["Unit price"].apply(clean_price) * 2,  # Moltiplica per 2
        "Taglia": df["Size US"].apply(format_taglia),  # Formatta la colonna Taglia
        "Barcode": df["EAN code"],  # Tratta il barcode come stringa
        "EAN": "",  # Colonna vuota aggiunta subito dopo Barcode
        "Qta": df["Quantity"],
        "Tot Costo": df["Unit price"].apply(clean_price) * df["Quantity"],  # Calcola Tot Costo come Costo * Qta
        "Materiale": "",
        "Spec. Materiale": "",
        "Misure": "",
        "Scala Taglie": "US",
        "Tacco": "",
        "Suola": "",
        "Carryover": "",
        "HS Code": ""
    })
    
    # Applica la funzione per duplicare le righe in base al valore di Qta
    expanded_df = expand_rows(output_df)
    
    return expanded_df

# Funzione per suddividere i dati in fogli di massimo 50 righe e aggiungere l'intestazione
def write_data_in_chunks(writer, df, stagione, data_inizio, data_fine, ricarico):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)  # Calcola il numero di fogli necessari
    for i in range(num_chunks):
        chunk_df = df[i*50:(i+1)*50]  # Estrai un blocco di massimo 50 righe

        # Scrivi i dati del DataFrame senza specificare sheet_name
        start_row = 9  # Riga in cui iniziano i dati (10 per l'utente)
        chunk_df.to_excel(writer, startrow=start_row, index=False)

        # Recupera l'oggetto del foglio corrente
        worksheet = writer.book.get_worksheet_by_name(writer.sheets.keys()[i])

        # Scrivi l'intestazione fissa nelle prime righe
        worksheet.write('A1', 'STAGIONE:')
        worksheet.write('B1', f"'{stagione}")  # Inserisci il valore di STAGIONE con apice
        worksheet.write('A2', 'TIPO:')
        worksheet.write('B2', 'ACCESSORI')  # Inserisci ACCESSORI accanto a TIPO
        worksheet.write('A3', 'DATA INIZIO:')
        worksheet.write('B3', f"'{data_inizio.strftime('%d/%m/%Y')}")  # Formatta correttamente la DATA INIZIO con apice
        worksheet.write('A4', 'DATA FINE:')
        worksheet.write('B4', f"'{data_fine.strftime('%d/%m/%Y')}")  # Formatta correttamente la DATA FINE con apice
        worksheet.write('A5', 'RICARICO:')
        worksheet.write('B5', ricarico)  # Inserisci il valore di RICARICO

        # Imposta la colonna "Barcode" come testo per evitare la notazione scientifica
        text_format = writer.book.add_format({'num_format': '@'})  # Formato per trattare come testo
        worksheet.set_column('L:L', 20, text_format)  # Colonna Barcode come testo

        # Trova l'ultima riga dei dati
        last_data_row = len(chunk_df) + start_row

        # Lascia una riga vuota
        empty_row = last_data_row + 1  # Riga vuota subito dopo l'ultima riga di dati
        worksheet.write(f'N{empty_row}', "")  # Riga vuota per Qta
        worksheet.write(f'O{empty_row}', "")  # Riga vuota per Tot Costo

        # Aggiungi la somma subito sotto la riga vuota
        total_row = empty_row + 2  # Riga per le somme, aggiungiamo +2 per lasciare una riga bianca
        number_format = writer.book.add_format({'num_format': '#,##0.00'})  # Formato numerico con due decimali
        worksheet.write_formula(f'N{total_row}', f"=SUM(N{start_row+2}:N{last_data_row + 1})", number_format)  # Somma per la colonna Qta
        worksheet.write_formula(f'O{total_row}', f"=SUM(O{start_row+2}:O{last_data_row + 1})", number_format)  # Somma per la colonna Tot Costo

        # Applica la formattazione numerica con due decimali alle colonne Qta e Tot Costo
        worksheet.set_column('N:N', None, number_format)
        worksheet.set_column('O:O', None, number_format)


# Streamlit app e scrittura del file
st.title('Asics Xmag')

# Campi di input per l'intestazione
stagione = st.text_input("Inserisci STAGIONE")
data_inizio = st.date_input("Inserisci DATA INIZIO")
data_fine = st.date_input("Inserisci DATA FINE")
ricarico = st.text_input("Inserisci RICARICO")

# Carica il file color.txt dalla directory del progetto
colors_mapping = load_colors_mapping("color.txt")  # Usa il percorso relativo alla main folder

# Permetti l'upload di più file Excel
uploaded_files = st.file_uploader("Scegli i file Excel", accept_multiple_files=True)

if uploaded_files and stagione and data_inizio and data_fine and ricarico:
    processed_dfs = []
    
    for uploaded_file in uploaded_files:
        processed_dfs.append(process_file(uploaded_file, colors_mapping))
    
    # Concatenate tutti i DataFrame
    final_df = pd.concat(processed_dfs, ignore_index=True)

    # Seleziona solo le combinazioni uniche di "Articolo" e "Colore"
    unique_combinations = final_df[["Articolo", "Colore"]].drop_duplicates()

    st.write("Anteprima Articolo-Colore:")

    # Dizionario per raccogliere il flag UOMO/DONNA per ogni combinazione
    selections = {}

    # Visualizzare l'anteprima dei dati unici con opzione per UOMO/DONNA
    for index, row in unique_combinations.iterrows():
        flag = st.selectbox(f"{row['Articolo']}-{row['Colore']}", options=["Seleziona...", "UOMO", "DONNA"], key=index)  # Selezione esplicita
        selections[(row['Articolo'], row['Colore'])] = flag

    if st.button("Elabora File"):
        # Controlla se tutte le selezioni sono valide (UOMO o DONNA)
        if any(flag == "Seleziona..." for flag in selections.values()):
            st.error("Devi selezionare UOMO o DONNA per tutte le combinazioni!")
        else:
            # Filtra i dati in base alla selezione UOMO/DONNA
            uomo_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UOMO', axis=1)]
            donna_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'DONNA', axis=1)]

            # Crea due file Excel distinti, uno per UOMO e uno per DONNA
            uomo_output = io.BytesIO()
            donna_output = io.BytesIO()

            # Scrivi i dati di UOMO con la logica di suddivisione in blocchi di 50 righe e aggiungi l'intestazione
            with pd.ExcelWriter(uomo_output, engine='xlsxwriter') as writer_uomo:
                write_data_in_chunks(writer_uomo, uomo_df, stagione, data_inizio, data_fine, ricarico)

            # Scrivi i dati di DONNA con la logica di suddivisione in blocchi di 50 righe e aggiungi l'intestazione
            with pd.ExcelWriter(donna_output, engine='xlsxwriter') as writer_donna:
                write_data_in_chunks(writer_donna, donna_df, stagione, data_inizio, data_fine, ricarico)

            # Fornisci due pulsanti separati per scaricare i file UOMO e DONNA
            st.download_button(
                label="Download File UOMO",
                data=uomo_output.getvalue(),
                file_name="uomo_processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.download_button(
                label="Download File DONNA",
                data=donna_output.getvalue(),
                file_name="donna_processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
