import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# Inizializza lo stato della sessione
reset_keys = [
    "Nome", "Secondo Nome", "Cognome", "Secondo Cognome", "Numero di Telefono", "Description", "Codice Fiscale",
    "Data di Fine", "Employee ID", "Dipartimento", "Email", "flag_email"
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
def genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno):
    nome = nome.split()[0]
    cognome = cognome.split()[0]
    secondo_nome = secondo_nome.split()[0] if secondo_nome else ""
    secondo_cognome = secondo_cognome.split()[0] if secondo_cognome else ""

    base = f"{nome[0].lower()}{secondo_nome[0].lower()}.{cognome.lower()}{secondo_cognome.lower()}"

    if esterno:
        limite = 16
        if len(base) > limite:
            base = f"{nome[0].lower()}{secondo_nome[0].lower()}.{cognome.lower()}"
        base += ".ext"
    else:
        limite = 20
        if len(base) > limite:
            base = f"{nome[0].lower()}{secondo_nome[0].lower()}.{cognome.lower()}"

    return base[:20]

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
    tipo_utente = st.selectbox("Seleziona il tipo di utente:", ["Dipendente Consip", "Esterno"])

    nome = st.text_input("Nome", key="Nome").strip().capitalize()
    secondo_nome = st.text_input("Secondo Nome", key="Secondo Nome").strip().capitalize()
    cognome = st.text_input("Cognome", key="Cognome").strip().capitalize()
    secondo_cognome = st.text_input("Secondo Cognome", key="Secondo Cognome").strip().capitalize()
    numero_telefono = st.text_input("Numero di Telefono", "", key="Numero di Telefono").replace(" ", "")
    description_input = st.text_input("Description (lascia vuoto per <PC>)", "<PC>", key="Description").strip()
    codice_fiscale = st.text_input("Codice Fiscale", "", key="Codice Fiscale").strip()

    expire_date = st.text_input("Data di Fine (gg-mm-aaaa)", "30-06-2025", key="Data di Fine")

    ou = st.selectbox("OU", ["Utenti standard", "Utenti VIP", "Utenti esterni - Consulenti", "Utenti esterni - Somministrati e Stage"])
    dipendente = st.selectbox("Tipo di Esterno:", ["", "Consulente", "Somministrato/Stage"])
    employee_id = st.text_input("Employee ID", "", key="Employee ID").strip()
    department = st.text_input("Dipartimento", "", key="Dipartimento").strip()

    inserimento_gruppo_default = "consip_vpn;dipendenti_wifi;mobile_wifi;GEDOGA-P-DOCGAR;GRPFreeDeskUser"
    inserimento_gruppo = inserimento_gruppo_default if tipo_utente == "Dipendente Consip" else "consip_vpn"

    telephone_number = "+39 06 854491" if tipo_utente == "Dipendente Consip" else ""
    company = "Consip" if tipo_utente == "Dipendente Consip" else ""

    email_flag = st.radio("Email necessaria? [Solo per Consulenti Esterni]", ["Sì", "No"], index=0, key="flag_email") == "Sì"
    email = st.text_input("Email Personalizzata (solo se Email = No)", "", key="Email").strip()

    employee_number = codice_fiscale

    if st.button("Genera CSV"):
        esterno = tipo_utente == "Esterno"
        sAMAccountName = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno)

        nome_completo = f"{nome} {secondo_nome} {cognome} {secondo_cognome}".strip()
        nome_completo = ' '.join(nome_completo.split())

        display_name = f"{nome_completo} (esterno)" if esterno else nome_completo
        expire_date_formatted = formatta_data(expire_date) if esterno else ""
        userprincipalname = f"{sAMAccountName}@consip.it"
        mobile = f"+39 {numero_telefono}" if numero_telefono else ""
        description = description_input if description_input else "<PC>"

        # Logica Email
        if tipo_utente == "Esterno" and dipendente == "Consulente" and not email_flag:
            mail = email
        else:
            mail = f"{sAMAccountName}@consip.it"

        given_name = f"{nome} {secondo_nome}".strip()
        surname = f"{cognome} {secondo_cognome}".strip()

        disable = "FALSE"
        moveToOU = ""
        
        row = [
            sAMAccountName,
            "TRUE",
            ou,
            nome_completo,
            display_name,
            nome_completo,
            given_name,
            surname,
            employee_number,
            employee_id,
            department,
            description,
            "TRUE",
            expire_date_formatted,
            userprincipalname,
            mail,
            mobile,
            "",
            inserimento_gruppo,
            disable,
            moveToOU,
            telephone_number,
            company
        ]

        # Crea un DataFrame
        df = pd.DataFrame([row], columns=header_modifica)

        # Scrive su un CSV in memoria
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)

        st.success("✅ CSV generato con successo!")
        st.download_button(
            label="📥 Scarica CSV",
            data=csv_buffer.getvalue(),
            file_name=f"{sAMAccountName}_output.csv",
            mime="text/csv"
        )

 
