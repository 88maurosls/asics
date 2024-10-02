import streamlit as st
import pandas as pd
import io
import os

# Funzione per formattare la colonna Taglia
def format_taglia(size_us):
    size_str = str(size_us)
    if ".0" in size_str:
        size_str = size_str.replace(".0", "")  # Rimuovi il .0
    return size_str.replace(".5", "+")  # Converti .5 in +

# Funzione per pulire i prezzi (rimuovi simbolo dell'euro e converti in float)
def clean_price(price):
    return float(str(price).replace("€", "").replace(",", "").strip())

# Funzione per leggere il file gender.txt e creare un dizionario di flag
def load_gender_data():
    gender_dict = {}
    if os.path.exists('gender.txt'):
        with open('gender.txt', 'r') as f:
            for line in f:
                parts = line.strip().split(',')
                if len(parts) == 3:  # Verifica che la riga contenga esattamente 3 elementi
                    articolo, colore, flag = parts
                    gender_dict[(articolo, colore)] = flag
    return gender_dict

# Funzione per salvare i flag selezionati in gender.txt
def save_gender_data(gender_dict):
    try:
        with open('gender.txt', 'w') as f:
            for (articolo, colore), flag in gender_dict.items():
                f.write(f"{articolo},{colore},{flag}\n")
        st.success("Dati salvati correttamente in gender.txt")
    except Exception as e:
        st.error(f"Errore durante il salvataggio: {e}")

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

# Carica i dati di genere esistenti da gender.txt
gender_data = load_gender_data()

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
        # Se la combinazione è già presente in gender.txt, seleziona automaticamente il flag
        default_flag = gender_data.get((row['Articolo'], row['Colore']), 'UOMO')
        flag = st.radio(f"Articolo: {row['Articolo']} - Colore: {row['Colore']}", ('UOMO', 'DONNA'), index=default_flag == 'DONNA')
        selections[(row['Articolo'], row['Colore'])] = flag

    if st.button("Elabora File"):
        # Filtra i dati in base alla selezione UOMO/DONNA
        uomo_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'UOMO', axis=1)]
        donna_df = final_df[final_df.apply(lambda x: selections[(x['Articolo'], x['Colore'])] == 'DONNA', axis=1)]

        # Salva i dati di genere aggiornati in gender.txt
        gender_data.update(selections)
        save_gender_data(gender_data)  # Salva le nuove selezioni nel file gender.txt

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
