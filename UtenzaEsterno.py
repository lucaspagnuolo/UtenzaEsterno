import streamlit as st
import csv
from datetime import datetime, timedelta

# Funzione per formattare la data (da gg-mm-aaaa a "mm/dd/aaaa 00:00") con gestione del cambio mese/anno
def formatta_data(data):
    giorno, mese, anno = map(int, data.split("-"))
    data_fine = datetime(anno, mese, giorno) + timedelta(days=1)
    return data_fine.strftime("%m/%d/%Y 00:00")

# Interfaccia Streamlit
st.title("Gestione Utenti Consip")

tipo_utente = st.selectbox("Seleziona il tipo di utente:", ["Dipendente Consip", "Esterno"])

nome = st.text_input("Nome").strip().capitalize()
cognome = st.text_input("Cognome").strip().capitalize()
numero_telefono = st.text_input("Numero di Telefono", "").replace(" ", "")
expire_date = st.text_input("Data di Fine (gg-mm-aaaa)", "30-06-2025")

if tipo_utente == "Dipendente Consip":
    ou = st.selectbox("OU", ["Utenti standard", "Utenti VIP"])
    employee_number = st.text_input("Codice Fiscale (Employee Number)", "").strip()
    employee_id = st.text_input("Employee ID", "").strip()
    department = st.text_input("Dipartimento", "").strip()
    inserimento_gruppo = "consip_vpn;dipendenti_wifi;mobile_wifi;GEDOGA-P-DOCGAR;GRPFreeDeskUser"
    telephone_number = "+39 06 854491"
    company = "Consip"
else:
    dipendente = st.selectbox("Tipo di Esterno:", ["Consulente", "Somministrato/Stage"])
    ou = "Utenti esterni - Consulenti" if dipendente == "Consulente" else "Utenti esterni - Somministrati e Stage"
    employee_number = ""
    employee_id = ""
    department = "Utente esterno"
    inserimento_gruppo = "consip_vpn"
    telephone_number = ""
    company = ""

if st.button("Genera CSV"):
    # Creazione valori comuni
    sAMAccountName = f"{nome.lower()}.{cognome.lower()}" if tipo_utente == "Dipendente Consip" else f"{nome.lower()}.{cognome.lower()}.ext"
    display_name = f"{cognome} {nome} (esterno)" if tipo_utente == "Esterno" else f"{cognome} {nome}"
    cn = display_name
    expire_date_formatted = formatta_data(expire_date)
    userprincipalname = f"{sAMAccountName}@consip.it"
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""

    # Riga CSV
    row = [
        sAMAccountName, "SI", ou, sAMAccountName, display_name, cn, nome, cognome,
        employee_number, employee_id, department, "<PC>", "No", expire_date_formatted,
        userprincipalname, userprincipalname, mobile, "", inserimento_gruppo, "", "", telephone_number, company
    ]

    # Creazione file CSV
    nome_file = f"{cognome}_{nome[0]}.csv"
    with open(nome_file, "w", newline="") as file:
        writer = csv.writer(file, quoting=csv.QUOTE_MINIMAL)
        writer.writerow([
            "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
            "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
            "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
            "disable", "moveToOU", "telephoneNumber", "company"
        ])
        writer.writerow(row)

    st.success(f"File CSV generato correttamente come '{nome_file}' con data di scadenza '{expire_date_formatted}'")
