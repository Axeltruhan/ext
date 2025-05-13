import streamlit as st
import pandas as pd
import random
from io import BytesIO
import tempfile
import os
import xlsxwriter

# Configure page settings
st.set_page_config(
    page_title="Estrazione righe", 
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items=None  # Nascondi menu
)

def filtra_e_seleziona(df):
    """
    Filtra e seleziona righe dal dataframe in base ai criteri:
    - Stato = 'Chiuso'
    - Sorgente = 'Web'
    - Iterazioni > 2
    - Per ogni gruppo in 'Assegnazione':
      - 1 riga con Processo = 'Change'
      - 5 righe con Processo != 'Change'
    - Tutti i "Motivo di Contatto" devono essere diversi tra loro
    - Le righe con "Modifica o correzione dati intestatario dominio e database" vanno sempre in fondo
    """
    gruppi = df['Assegnazione'].unique()
    risultati = []

    for gruppo in gruppi:
        gruppo_df = df[
            (df['Assegnazione'] == gruppo) &
            (df['Stato'] == 'Chiuso') &
            (df['Sorgente'] == 'Web') &
            (df['Iterazioni'] > 2)
        ]

        change_rows = gruppo_df[gruppo_df['Processo'] == 'Change']
        non_change_rows = gruppo_df[gruppo_df['Processo'] != 'Change']

        if len(change_rows) < 1 or len(non_change_rows) < 5:
            continue  # Skip se non ci sono abbastanza dati validi
            
        # Dividiamo le righe con "Modifica o correzione dati" dalle altre
        modifica_rows = pd.DataFrame()
        other_rows = pd.DataFrame()
        
        # Filtriamo le righe con "Modifica o correzione dati intestatario dominio e database"
        modifica_pattern = "Modifica o correzione dati intestatario dominio e database"
        
        # Selezioniamo il campione assicurandoci che i "Motivo di Contatto" siano diversi
        selected_rows = []
        selected_motivi = set()
        
        # Prima prendiamo una riga 'Change'
        change_sample = None
        for _, row in change_rows.iterrows():
            motivo = row['Motivo di Contatto']
            if motivo not in selected_motivi:
                selected_motivi.add(motivo)
                change_sample = row.to_frame().T
                break
        
        # Se non troviamo un motivo unico, prendiamo semplicemente il primo
        if change_sample is None and len(change_rows) > 0:
            change_sample = change_rows.iloc[[0]]
            selected_motivi.add(change_sample.iloc[0]['Motivo di Contatto'])
        
        # Ora prendiamo 5 righe 'non-Change' con motivi diversi
        non_change_sample = pd.DataFrame()
        count = 0
        
        # Separiamo le righe con "Modifica o correzione dati"
        modifica_rows_nonchange = non_change_rows[non_change_rows['Motivo di Contatto'].str.contains(modifica_pattern, case=False, na=False)]
        other_rows_nonchange = non_change_rows[~non_change_rows['Motivo di Contatto'].str.contains(modifica_pattern, case=False, na=False)]
        
        # Prima prendiamo righe non "Modifica o correzione dati"
        for _, row in other_rows_nonchange.iterrows():
            motivo = row['Motivo di Contatto']
            if motivo not in selected_motivi and count < 5:
                selected_motivi.add(motivo)
                non_change_sample = pd.concat([non_change_sample, row.to_frame().T])
                count += 1
                
        # Se abbiamo ancora bisogno di righe, usiamo quelle "Modifica o correzione dati"
        for _, row in modifica_rows_nonchange.iterrows():
            motivo = row['Motivo di Contatto']
            if motivo not in selected_motivi and count < 5:
                selected_motivi.add(motivo)
                non_change_sample = pd.concat([non_change_sample, row.to_frame().T])
                count += 1
        
        # Se non abbiamo abbastanza motivi diversi, prendiamo i rimanenti anche se duplicati
        if count < 5 and len(non_change_rows) >= 5:
            remaining = 5 - count
            remaining_rows = non_change_rows[~non_change_rows.index.isin(non_change_sample.index)].sample(min(remaining, len(non_change_rows) - count))
            non_change_sample = pd.concat([non_change_sample, remaining_rows])
        
        # Se abbiamo sia Change che non-Change, procediamo
        if change_sample is not None and len(non_change_sample) > 0:
            # Ordiniamo il campione mettendo le righe "Modifica o correzione dati" in fondo
            sample = pd.concat([change_sample, non_change_sample])
            
            # Riordina mettendo "Modifica o correzione dati" in fondo
            has_modifica = sample['Motivo di Contatto'].str.contains(modifica_pattern, case=False, na=False)
            ordered_sample = pd.concat([sample[~has_modifica], sample[has_modifica]])
            
            risultati.append(ordered_sample)

    if risultati:
        # Concatena tutti i risultati
        result_df = pd.concat(risultati)
        
        # Ordina il DataFrame per 'Assegnazione' in ordine alfabetico
        # e mantiene l'ordinamento all'interno di ogni gruppo di 'Assegnazione'
        # (prima le non-modifica, poi le modifica)
        
        # Prima crea una colonna temporanea che indica se è una riga "Modifica o correzione dati"
        modifica_pattern = "Modifica o correzione dati intestatario dominio e database"
        result_df['_is_modifica'] = result_df['Motivo di Contatto'].str.contains(modifica_pattern, case=False, na=False)
        
        # Ora ordina per Assegnazione (alfabetico) e poi per _is_modifica
        # In questo modo manteniamo l'ordine: prima tutte le assegnazioni in ordine alfabetico,
        # e all'interno di ogni assegnazione, prima le non-modifica e poi le modifica
        result_df = result_df.sort_values(by=['Assegnazione', '_is_modifica'])
        
        # Rimuove la colonna temporanea
        result_df = result_df.drop(columns=['_is_modifica'])
        
        return result_df
    else:
        return pd.DataFrame()

def get_excel_download(df):
    """
    Crea un file Excel (.xlsx) utilizzando openpyxl - versione semplificata
    """
    # Creiamo un nome file temporaneo (ma non lo useremo come file)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmpfile:
        temp_path = tmpfile.name
    
    # Salviamo il file Excel in modo normale senza BytesIO per evitare conflitti
    # Usiamo direttamente openpyxl che è più affidabile di xlsxwriter in alcuni casi
    df.to_excel(temp_path, sheet_name='Risultati', index=False, engine='openpyxl')
    
    # Leggiamo il file come binario
    with open(temp_path, 'rb') as f:
        data = f.read()
    
    # Rimuoviamo il file temporaneo
    os.unlink(temp_path)
    
    return data

def main():
    """
    Funzione principale dell'applicazione
    """
    # Inizializzazione dello stato della sessione per mantenere i dati
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'risultato' not in st.session_state:
        st.session_state.risultato = pd.DataFrame()
    if 'file_caricato' not in st.session_state:
        st.session_state.file_caricato = False
        
    # Header semplificato con titolo centrato
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h1 style='text-align: center;'>🎯 Estrazione righe</h1>", unsafe_allow_html=True)
    
    # CSS Generale per tutta l'app
    st.markdown("""
    <style>
    /* Stile globale */
    button[kind="secondary"] {
        box-shadow: none !important;
        background-color: #f0f2f6 !important;
    }
    /* Rimuovi la barra grigia a fianco dei pulsanti */
    div.stButton > button {
        width: auto !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: none !important;
        border-radius: 4px !important;
    }
    div.row-widget.stButton {
        width: auto !important;
        background-color: transparent !important;
    }

    
    /* Per lo stile dei componenti */
    .uploadedFile {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Solo pulsante per caricamento del file Excel centrato
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # File uploader standard con label in italiano
        uploaded_file = st.file_uploader(
            "Carica il file Excel", 
            type=['xlsx'],
            help="Trascina e rilascia il file Excel qui o clicca per cercarlo nel tuo computer"
        )
    
    # Se un nuovo file è stato caricato, aggiorniamo il dataframe nello stato della sessione
    if uploaded_file is not None and (st.session_state.df is None or not st.session_state.file_caricato):
        try:
            st.session_state.df = pd.read_excel(uploaded_file)
            st.session_state.file_caricato = True
            
            # Controllo colonne necessarie
            colonne_necessarie = {'Assegnazione', 'Stato', 'Sorgente', 'Iterazioni', 'Processo', 'ID', 'Motivo di Contatto'}
            if not colonne_necessarie.issubset(set(st.session_state.df.columns)):
                st.error(f"Il file deve contenere le colonne: {', '.join(colonne_necessarie)}")
                st.session_state.file_caricato = False
                return
            
            # Info sul file caricato
            st.success(f"File caricato con successo! ({st.session_state.df.shape[0]} righe, {st.session_state.df.shape[1]} colonne)")
            
            with st.expander("📊 Informazioni sul dataset"):
                st.write(f"**Valori unici in 'Assegnazione':** {len(st.session_state.df['Assegnazione'].unique())}")
                st.write(f"**Totale ticket con Stato 'Chiuso':** {len(st.session_state.df[st.session_state.df['Stato'] == 'Chiuso'])}")
                st.write(f"**Totale ticket con Sorgente 'Web':** {len(st.session_state.df[st.session_state.df['Sorgente'] == 'Web'])}")
                st.write(f"**Totale ticket con Iterazioni > 2:** {len(st.session_state.df[st.session_state.df['Iterazioni'] > 2])}")
                st.write(f"**Totale ticket con Processo 'Change':** {len(st.session_state.df[st.session_state.df['Processo'] == 'Change'])}")
        except Exception as e:
            st.error(f"Si è verificato un errore durante il caricamento del file: {str(e)}")
            return
    
    # Se abbiamo un file caricato, mostriamo il pulsante per estrarre
    if st.session_state.file_caricato:
        # Pulsante per elaborare i dati
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            # Utilizzare un componente HTML personalizzato per il pulsante
            st.markdown('<div style="text-align: center; margin-bottom: 20px;">', unsafe_allow_html=True)
            pulsante_estrai = st.button("🔍 Estrai", use_container_width=False)
            
        if pulsante_estrai:
            with st.spinner("Elaborazione in corso..."):
                st.session_state.risultato = filtra_e_seleziona(st.session_state.df)

            
        # Mostriamo i risultati se disponibili
        if not st.session_state.risultato.empty:
            st.header("✅ Risultati")
            st.success(f"Campione selezionato con successo! ({len(st.session_state.risultato)} ticket)")
            
            # Preparazione della tabella per la visualizzazione
            tabella_visuale = st.session_state.risultato[['Iterazioni', 'ID', 'Motivo di Contatto', 'Assegnazione']].copy()
            
            # Configura la visualizzazione con colonne direttamente copiabili
            column_config = {
                'Iterazioni': st.column_config.NumberColumn('Iterazioni'),
                'ID': st.column_config.TextColumn(
                    'ID',
                    help='Clicca per copiare il valore',
                    disabled=False
                ),
                'Motivo di Contatto': st.column_config.TextColumn(
                    'Motivo di Contatto',
                    help='Clicca per copiare il valore',
                    disabled=False
                ),
                'Assegnazione': st.column_config.TextColumn('Assegnazione')
            }
            
            # Evidenziamo le righe con "Modifica o correzione dati intestatario dominio e database"
            modifica_pattern = "Modifica o correzione dati intestatario dominio e database"
            
            # Identifichiamo le righe che contengono il pattern
            mask_modifica = tabella_visuale['Motivo di Contatto'].str.contains(modifica_pattern, case=False, na=False)
            
            # Aggiungiamo informazioni per l'utente su quali righe sono evidenziate
            modifica_rows = mask_modifica.sum()
            if modifica_rows > 0:
                st.info(f"Ci sono {modifica_rows} righe con 'Modifica o correzione dati intestatario dominio e database' posizionate in fondo e evidenziate con sfondo giallo paglierino.")
            
            # Creiamo una versione stilizzata del dataframe
            styled_df = tabella_visuale.style.apply(
                lambda x: ['background-color: #FFF9C4' if mask_modifica.iloc[i] else '' 
                          for i in range(len(x))], 
                axis=1
            )
            
            # Mostra la tabella con lo styling condizionale
            st.dataframe(
                styled_df,
                use_container_width=True,
                column_config=column_config,
                hide_index=True
            )
            
            # Aggiungiamo pulsanti per copiare intere colonne
            st.write("##### Copia colonne:")
            
            # Organizziamo i pulsanti in colonne
            cols = st.columns(3)
            
            # Definiamo le colonne da copiare con pulsanti
            colonne_copiabili = ['ID', 'Motivo di Contatto', 'Assegnazione']
            
            # Funzione per generare contenuto della colonna
            def get_column_text(col_name):
                return '\n'.join(tabella_visuale[col_name].astype(str).tolist())
                
            # Creiamo aree di testo espandibili per ogni colonna copiabile
            for i, col_name in enumerate(colonne_copiabili):
                with cols[i]:
                    # Prepariamo i dati da copiare (tutti i valori della colonna)
                    values = get_column_text(col_name)
                    
                    # Aggiungiamo expander per ogni colonna
                    with st.expander(f"📋 Copia tutti '{col_name}'", expanded=False):
                        # Area di testo con i valori della colonna
                        st.code(values, language="text")
                        st.caption("Seleziona tutto il testo sopra e copialo con Ctrl+C / Cmd+C")
            

        elif pulsante_estrai:  # Solo se è stato premuto il pulsante e non ci sono risultati
            st.warning("⚠️ Nessun gruppo valido trovato con i criteri indicati.")
            st.info("""
            Per ogni gruppo in 'Assegnazione', sono necessari:
            - Almeno 1 ticket con Processo = 'Change'
            - Almeno 5 ticket con Processo ≠ 'Change'
            - Tutti i ticket devono avere Stato = 'Chiuso', Sorgente = 'Web' e Iterazioni > 2
            """)

if __name__ == "__main__":
    main()