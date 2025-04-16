import streamlit as st
import csv
from datetime import datetime

st.title("Generatore Account AD")

# Input utente
nome = st.text_input("Nome").strip()
cognome = st.text_input("Cognome").strip()
tipo_utente = st.selectbox("Tipologia Utente", ["Interno", "Esterno"])
dipendente = ""
if tipo_utente == "Esterno":
    dipendente = st.selectbox("Qualifica", ["Fornitore", "Consulente", "Altro"])

azienda = st.text_input("Azienda", "").strip()
azienda_codice = st.text_input("Codice Fiscale Azienda", "").strip()
company = azienda if azienda else "Consip"

# Nome completo e sAMAccountName
nome_completo = f"{cognome} {nome}"
sAMAccountName = f"{nome[0].lower()}{cognome.lower()}"
display_name = nome_completo
given_name = nome
surname = cognome

# Matricola o identificativo esterno
if tipo_utente == "Interno":
    employee_number = st.text_input("Matricola").strip()
    employee_id = employee_number
else:
    employee_number = st.text_input("Identificativo Esterno").strip()
    employee_id = employee_number

# Department e Description
department = st.text_input("Unità Organizzativa", "").strip()
description = st.text_input("Ruolo/Descrizione", "").strip()

# Data di scadenza account
expire_date = st.date_input("Data di Scadenza Account")
expire_date_formatted = expire_date.strftime("%d/%m/%Y")

# Cellulare e telefono
mobile = st.text_input("Cellulare", "").strip()
telephone_number = st.text_input("Telefono Fisso", "").strip()

# Username (UPN)
userprincipalname = f"{sAMAccountName}@consip.it"

# Email
email = ""
if tipo_utente == "Esterno" and dipendente == "Consulente":
    email_flag = st.radio("Email necessaria?", ["Sì", "No"], index=0, key="flag_email") == "Sì"
else:
    email_flag = True  # sempre sì per tutti gli altri

# Generazione email solo se necessario
if email_flag:
    email = f"{sAMAccountName}@consip.it"

# Inserimento Gruppo (es. O365 Utente Standard)
inserimento_gruppo = st.text_input("Gruppo da inserire", "").strip()

# OU
ou_default = "OU=Users,DC=consip,DC=it"
ou = st.text_input("Organizational Unit (OU)", ou_default).strip()

# Pulsante per generare
if st.button("Genera CSV"):
    # Prepara riga
    row = [
        sAMAccountName, "SI", ou, nome_completo, display_name, nome_completo, given_name, surname,
        employee_number, employee_id, department, description, "No", expire_date_formatted,
        userprincipalname, email, mobile, "", inserimento_gruppo, "", "", telephone_number, company
    ]

    # Scrive su file
    filename = f"output_account_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow([
            "sAMAccountName", "abilita", "OU", "nome completo", "displayName", "cn", "givenName", "sn",
            "employeeNumber", "employeeID", "department", "description", "changePasswordAtLogon",
            "accountExpires", "userPrincipalName", "mail", "mobile", "proxyAddresses",
            "memberof", "title", "physicalDeliveryOfficeName", "telephoneNumber", "company"
        ])
        writer.writerow(row)

    st.success(f"CSV generato: {filename}")
