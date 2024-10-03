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

# Funzione per duplicare le righe in base al valore di Qta
def expand_rows(df):
    return df.loc[df.index.repeat(df['Qta'])].assign(Qta=1)

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
        "EAN": "",  # Colonna vuota aggiunta subito dopo Barcode
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
    
    # Applica la funzione per duplicare le righe in base al valore di Qta
    expanded_df = expand_rows(output_df)
    
    return expanded_df

# Funzione per suddividere i dati in fogli di massimo 50 righe
def write_data_in_chunks(writer, df, sheet_name_base):
    num_chunks = len(df) // 50 + (1 if len(df) % 50 > 0 else 0)  # Calcola il numero di fogli necessari
    for i in range(num_chunks):
        chunk_df = df[i*50:(i+1)*50]  # Estrai un blocco di massimo 50 righe
        sheet_name = f"{sheet_name_base}" if i == 0 else f"{sheet_name_base} {i+1}"  # Nome del foglio (UOMO, UOMO 2, etc.)
        chunk_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Imposta la colonna "Barcode" come testo per evitare la notazione scientifica
        worksheet = writer.sheets[sheet_name]
        text_format = writer.book.add_format({'num_format': '@'})  # Formato per trattare come testo
        worksheet.set_column('L:L', 20, text_format)  # Formatta la colonna Barcode come testo

# Streamlit app e scrittura del file
st.title('Asics Xmag')

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

            # Crea un file in memoria per il download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Scrivi i dati in blocchi di 50 righe separati in fogli distinti
                write_data_in_chunks(writer, uomo_df, 'UOMO')
                write_data_in_chunks(writer, donna_df, 'DONNA')

            # Fornisci un pulsante per scaricare il file elaborato
            st.download_button(
                label="Download Processed Excel",
                data=output.getvalue(),
                file_name="processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
