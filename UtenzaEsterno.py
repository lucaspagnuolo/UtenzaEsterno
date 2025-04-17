import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# Inizializza lo stato della sessione
reset_keys = [
    "Nome", "Secondo Nome", "Cognome", "Secondo Cognome", "Numero di Telefono", "Description", "Codice Fiscale",
    "Data di Fine", "Employee ID", "Dipartimento", "Email", "flag_email", "Email Aziendale", "Manager", "Profilare sulla SM"
]

if "reset_fields" not in st.session_state:
    st.session_state.reset_fields = False

# Pulsante per pulire i campi
if st.button("🔄 Pulisci Campi"):
    st.session_state.reset_fields = True

# Funzione per formattare la data
def formatta_data(data):
    for separatore in ["-", "/"]:
        try:
            giorno, mese, anno = map(int, data.split(separatore))
            data_fine = datetime(anno, mese, giorno) + timedelta(days=1)
            return data_fine.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

# Funzione per generare sAMAccountName
def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False):
    # Pulizia e normalizzazione input
    nome = nome.strip().lower()
    secondo_nome = secondo_nome.strip().lower() if secondo_nome else ""
    cognome = cognome.strip().lower()
    secondo_cognome = secondo_cognome.strip().lower() if secondo_cognome else ""

    # Suffisso esterni
    suffix = ".ext" if esterno else ""
    limite = 16 if esterno else 20

    # Tentativo 1: nome + secondo_nome + . + cognome + secondo_cognome
    full_1 = f"{nome}{secondo_nome}.{cognome}{secondo_cognome}"
    if len(full_1) <= limite:
        return full_1 + suffix

    # Tentativo 2: iniziali + . + cognome + secondo_cognome
    iniz_nome = nome
    iniz_sec_nome = secondo_nome if secondo_nome else ""
    full_2 = f"{iniz_nome}{iniz_sec_nome}.{cognome}{secondo_cognome}"
    if len(full_2) <= limite:
        return full_2 + suffix

    # Tentativo 3: iniziali + . + cognome (senza secondo cognome)
    full_3 = f"{iniz_nome}{iniz_sec_nome}.{cognome}"
    return full_3[:limite] + suffix  # Garantisce che il limite non venga superato anche in casi estremi

# Se reset_fields è attivo, azzera i campi
if st.session_state.reset_fields:
    for key in reset_keys:
        st.session_state[key] = ""
    st.session_state.reset_fields = False

# Interfaccia
st.title("Gestione Utenze Consip")

funzionalita = st.radio("Scegli funzionalità:", ["Gestione Creazione Utenze", "Gestione Modifiche AD"])

header_modifica = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

if funzionalita == "Gestione Creazione Utenze":
    tipo_utente = st.selectbox("Seleziona il tipo di utente:", ["Dipendente Consip", "Esterno", "Azure"])

    nome = st.text_input("Nome", key="Nome").strip().capitalize()
    secondo_nome = st.text_input("Secondo Nome", key="Secondo Nome").strip().capitalize()
    cognome = st.text_input("Cognome", key="Cognome").strip().capitalize()
    secondo_cognome = st.text_input("Secondo Cognome", key="Secondo Cognome").strip().capitalize()
    numero_telefono = st.text_input("Numero di Telefono", "", key="Numero di Telefono").replace(" ", "")
    description_input = st.text_input("Description (lascia vuoto per <PC>)", "<PC>", key="Description").strip()

    if tipo_utente == "Azure":
        email_aziendale = st.text_input("Email Aziendale", key="Email Aziendale").strip()
        manager = st.text_input("Manager", key="Manager").strip()
        profila
