import streamlit as st
import subprocess
import os
import sys

st.set_page_config(page_title="OMNIA-TFS Gestione Schede", page_icon="📋", layout="centered")

st.title("📋 Sistema Gestione Schede OMNIA-TFS")
st.markdown("*Federazione Nazionale Maestri del Lavoro - Settore Scuola*")

tab1, tab2 = st.tabs(["1. Genera Nuove Schede", "2. Elabora Scansioni"])

with tab1:
    st.subheader("Configurazione e Generazione PDF")
    
    col1, col2 = st.columns(2)
    with col1:
        consolato = st.text_input("Consolato", value="MN")
        scuola = st.text_input("Scuola", value="ITIS Fermi")
        classe = st.text_input("Classe", value="3A")
    with col2:
        data = st.text_input("Data (AAAA-MM-GG)", value="2026-10-14")
        relatore = st.text_input("Relatore", value="Rossi")
        copie = st.number_input("Numero di Copie", min_value=1, value=25, step=1)

    if st.button("Genera PDF Schede", type="primary"):
        stringa_rif = f"{consolato}|{scuola}|{classe}|{data}|{relatore}"
        nome_pdf = f"schede_{classe.replace(' ', '_')}.pdf"
        
        with st.spinner("Generazione del PDF in corso..."):
            # Usiamo sys.executable per puntare al Python corretto del server
            cmd = [sys.executable, "genera_schede.py", "--rif", stringa_rif, "--copie", str(int(copie)), "--out", nome_pdf, "--template", "template.json"]
            subprocess.run(cmd, check=True)
            
        if os.path.exists(nome_pdf):
            st.success("PDF generato con successo!")
            with open(nome_pdf, "rb") as file:
                st.download_button(label="📥 Scarica il PDF delle Schede", data=file, file_name=nome_pdf, mime="application/pdf")

with tab2:
    st.subheader("Lettura delle schede scansionate")
    st.markdown("Carica il file PDF contenente tutte le schede compilate a mano e scansionate.")
    
    file_scansione = st.file_uploader("Scegli il file PDF delle scansioni", type=["pdf"])
    
    if file_scansione is not None:
        if st.button("Elabora e Leggi Schede", type="primary"):
            with open("temp_scansione.pdf", "wb") as f:
                f.write(file_scansione.getbuffer())
                
            nome_excel = "risultati_valutazione.xlsx"
            
            with st.spinner("Analisi ottica (OMR) in corso... Attendere."):
                cmd = [sys.executable, "leggi_schede.py", "temp_scansione.pdf", "--template", "template.json", "--out", nome_excel, "--commenti", "cartella_commenti"]
                result = subprocess.run(cmd, capture_output=True, text=True)
            
            if os.path.exists(nome_excel):
                st.success("Elaborazione completata!")
                with open(nome_excel, "rb") as file:
                    st.download_button(label="📥 Scarica i Risultati in Excel", data=file, file_name=nome_excel, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("Errore durante l'elaborazione delle schede.")
                st.code(result.stderr)
