import streamlit as st
import pandas as pd
from datetime import datetime
import io

def genera_samaccountname(nome, secondo_nome, cognome, secondo_cognome, esterno=False):
    # Rimuovi spazi e converti in minuscolo
    nome = nome.strip().lower() if nome else ''
    secondo_nome = secondo_nome.strip().lower() if secondo_nome else ''
    cognome = cognome.strip().lower() if cognome else ''
    secondo_cognome = secondo_cognome.strip().lower() if secondo_cognome else ''

    # Prendi le iniziali
    iniziale_nome = nome[0] if nome else ''
    iniziale_secondo_nome = secondo_nome[0] if secondo_nome else ''
    
    cognome_completo = cognome + secondo_cognome

    if esterno:
        sam = f"{iniziale_nome}{iniziale_secondo_nome}{cognome_completo}_est"
    else:
        sam = f"{iniziale_nome}{iniziale_secondo_nome}{cognome_completo}"

    return sam
st.title("Generatore CSV - Utenze Azure ed Esterne")

uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

rows = []

for _, row in df.iterrows():
    tipo_utenza = row.get("Tipo Utenza", "").strip().upper()
    nome = str(row.get("Nome", "")).strip()
    secondo_nome = str(row.get("Secondo Nome", "")).strip()
    cognome = str(row.get("Cognome", "")).strip()
    secondo_cognome = str(row.get("Secondo Cognome", "")).strip()
    cf = str(row.get("Codice Fiscale", "")).strip()
    mail = str(row.get("Email", "")).strip()
    descrizione = str(row.get("Descrizione", "")).strip()
    societa = str(row.get("Società", "")).strip()
    funzione = str(row.get("Funzione", "")).strip()
    referente = str(row.get("Referente", "")).strip()
    matricola = str(row.get("Matricola", "")).strip()

    # Elaborazione per utenze Azure
    if tipo_utenza == "AZURE":
        sam = genera_samaccountname(nome, secondo_nome, cognome, secondo_cognome, esterno=False)
        rows.append({
            "SamAccountName": sam,
            "Nome": nome,
            "Secondo Nome": secondo_nome,
            "Cognome": cognome,
            "Secondo Cognome": secondo_cognome,
            "Codice Fiscale": cf,
            "Email": mail,
            "Tipo Utenza": tipo_utenza,
            "Descrizione": descrizione,
            "Società": societa,
            "Funzione": funzione,
            "Referente": referente,
            "Matricola": matricola
        })

    # Elaborazione per utenze Esterne
    elif tipo_utenza == "ESTERNA":
        sam = genera_samaccountname(nome, secondo_nome, cognome, secondo_cognome, esterno=True)
        expire_date = row.get("Data Scadenza", None)

        if isinstance(expire_date, pd.Timestamp):
            expire_date = expire_date.strftime("%Y-%m-%d")
        elif expire_date is None or expire_date == "":
            expire_date = ""

        rows.append({
            "SamAccountName": sam,
            "Nome": nome,
            "Secondo Nome": secondo_nome,
            "Cognome": cognome,
            "Secondo Cognome": secondo_cognome,
            "Codice Fiscale": cf,
            "Email": mail,
            "Tipo Utenza": tipo_utenza,
            "Descrizione": descrizione,
            "Società": societa,
            "Funzione": funzione,
            "Referente": referente,
            "Matricola": matricola,
            "Data Scadenza": expire_date
        })

if rows:
    output_df = pd.DataFrame(rows)

    csv_buffer = io.StringIO()
    output_df.to_csv(csv_buffer, index=False, sep=";")
    csv_bytes = csv_buffer.getvalue().encode("utf-8")

    st.download_button(
        label="📥 Scarica CSV Generato",
        data=csv_bytes,
        file_name="utenze_generato.csv",
        mime="text/csv"
    )

    st.success("✅ CSV generato con successo!")
