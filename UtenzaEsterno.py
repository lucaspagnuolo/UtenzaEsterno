import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# Inizializza lo stato della sessione
reset_keys = [
    "Nome", "Cognome", "Numero di Telefono", "Description", "Codice Fiscale",
    "Data di Fine", "Employee ID", "Dipartimento"
]

if "reset_fields" not in st.session_state:
    st.session_state.reset_fields = False

# Pulsante per pulire i campi
if st.button("🔄 Pulisci Campi"):
    st.session_state.reset_fields = True

# Funzione per formattare la data
def formatta_data(data):
    giorno, mese, anno = map(int, data.split("-"))
    data_fine = datetime(anno, mese, giorno) + timedelta(days=1)
    return data_fine.strftime("%m/%d/%Y 00:00")

# Funzione per generare sAMAccountName
def genera_samaccountname(nome, cognome, esterno):
    nome = nome.split()[0]
    cognome = cognome.split()[0]
    base = f"{nome.lower()}.{cognome.lower()}"
    if esterno:
        base += ".ext"
    if len(base) > 20:
        base = f"{nome[0].lower()}.{cognome.lower()}"
        if esterno:
            base += ".ext"
    return base[:20]

# Se reset_fields è attivo, azzera i campi
if st.session_state.reset_fields:
    for key in reset_keys:
        st.session_state[key] = ""
    st.session_state.reset_fields = False

# Interfaccia
st.title("Creazioni Utenze Consip")

tipo_utente = st.selectbox("Seleziona il tipo di utente:", ["Dipendente Consip", "Esterno"])

nome = st.text_input("Nome", key="Nome").strip().capitalize()
cognome = st.text_input("Cognome", key="Cognome").strip().capitalize()
numero_telefono = st.text_input("Numero di Telefono", "", key="Numero di Telefono").replace(" ", "")
description_input = st.text_input("Description (lascia vuoto per <PC>)", "<PC>", key="Description").strip()
codice_fiscale = st.text_input("Codice Fiscale", "", key="Codice Fiscale").strip()

expire_date = ""
if tipo_utente == "Esterno":
    expire_date = st.text_input("Data di Fine (gg-mm-aaaa)", "30-06-2025", key="Data di Fine")

department = ""
if tipo_utente == "Dipendente Consip":
    ou = st.selectbox("OU", ["Utenti standard", "Utenti VIP"])
    employee_id = st.text_input("Employee ID", "", key="Employee ID").strip()
    department = st.text_input("Dipartimento", "", key="Dipartimento").strip()
    inserimento_gruppo = "consip_vpn;dipendenti_wifi;mobile_wifi;GEDOGA-P-DOCGAR;GRPFreeDeskUser"
    telephone_number = "+39 06 854491"
    company = "Consip"
else:
    dipendente = st.selectbox("Tipo di Esterno:", ["Consulente", "Somministrato/Stage"])
    ou = "Utenti esterni - Consulenti" if dipendente == "Consulente" else "Utenti esterni - Somministrati e Stage"
    employee_id = ""
    department = "Utente esterno"
    if dipendente == "Somministrato/Stage":
        inserimento_gruppo = "consip_vpn;dipendenti_wifi;mobile_wifi;GRPFreeDeskUser"
        department = st.text_input("Dipartimento", "", key="Dipartimento").strip()
    else:
        inserimento_gruppo = "consip_vpn"
    telephone_number = ""
    company = ""

employee_number = codice_fiscale

if st.button("Genera CSV"):
    esterno = tipo_utente == "Esterno"
    sAMAccountName = genera_samaccountname(nome, cognome, esterno)
    display_name = f"{cognome} {nome} (esterno)" if esterno else f"{cognome} {nome}"
    expire_date_formatted = formatta_data(expire_date) if esterno else ""
    userprincipalname = f"{sAMAccountName}@consip.it"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    description = description_input if description_input else "<PC>"

    # === CSV Principale ===
    row = [
        sAMAccountName, "SI", ou, display_name, display_name, display_name, nome, cognome,
        employee_number, employee_id, department, description, "No", expire_date_formatted,
        userprincipalname, userprincipalname, mobile, "", inserimento_gruppo, "", "", telephone_number, company
    ]

    header_main = [
        "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
        "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
        "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
        "disable", "moveToOU", "telephoneNumber", "company"
    ]

    output_main = io.StringIO()
    writer_main = csv.writer(output_main, quoting=csv.QUOTE_MINIMAL)
    writer_main.writerow(header_main)
    writer_main.writerow(row)
    output_main.seek(0)

    df_main = pd.DataFrame([row], columns=header_main)
    st.dataframe(df_main)

    st.download_button(
        label="📥 Scarica CSV Utente",
        data=output_main.getvalue(),
        file_name=f"{cognome}_{nome[0]}_utente.csv",
        mime="text/csv"
    )

    # === CSV Aggiuntivo ===
    row_extra = [
        description,  # Computer
        ou,           # OU
        userprincipalname,  # add_mail
        "",           # remove_mail
        mobile,       # add_mobile
        "",           # remove_mobile
        sAMAccountName,  # add_userprincipalname
        "",           # remove_userprincipalname
        "",           # disable
        ""            # moveToOU
    ]

    header_extra = [
        "Computer", "OU", "add_mail", "remove_mail", "add_mobile", "remove_mobile",
        "add_userprincipalname", "remove_userprincipalname", "disable", "moveToOU"
    ]

    output_extra = io.StringIO()
    writer_extra = csv.writer(output_extra, quoting=csv.QUOTE_MINIMAL)
    writer_extra.writerow(header_extra)
    writer_extra.writerow(row_extra)
    output_extra.seek(0)

    df_extra = pd.DataFrame([row_extra], columns=header_extra)
    st.dataframe(df_extra)

    st.download_button(
        label="📥 Scarica CSV Extra",
        data=output_extra.getvalue(),
        file_name=f"{cognome}_{nome[0]}_extra.csv",
        mime="text/csv"
    )

    st.success(f"✅ File CSV generati con successo per '{sAMAccountName}'")
