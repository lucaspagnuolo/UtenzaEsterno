import streamlit as st
import csv
from datetime import datetime, timedelta

# Funzione per formattare la data
def formatta_data(data):
    giorno, mese, anno = map(int, data.split("-"))
    data_fine = datetime(anno, mese, giorno) + timedelta(days=1)
    return data_fine.strftime("%m/%d/%Y 00:00")

# Interfaccia Streamlit
st.title("Generatore di CSV per Utenti Esterni")

# Input utente
dipendente = st.selectbox("Tipo di dipendente", ["consulente", "somministrato/stage"])
nome = st.text_input("Nome")
cognome = st.text_input("Cognome")
CF = st.text_input("Codice Fiscale")
numero_telefono = st.text_input("Numero di telefono")
data_fine = st.text_input("Data fine (formato: gg-mm-aaaa)")

if st.button("Genera CSV"):
    if not (nome and cognome and CF and numero_telefono and data_fine):
        st.error("Compila tutti i campi prima di generare il CSV.")
    else:
        # Creazione valori
        nome_cap = nome.split()[0].capitalize()
        cognome_cap = cognome.capitalize()
        sAMAccountName = f"{nome_cap.lower()}.{cognome_cap.lower()}.ext"
        OU = "Utenti esterni - Consulenti" if dipendente == "consulente" else "Utenti esterni - Somministrati e Stage"
        DisplayName = f"{cognome_cap} {nome_cap} (esterno)"
        cn = DisplayName
        ExpireDate = formatta_data(data_fine)
        userprincipalname = f"{sAMAccountName}@consip.it"
        mobile = f"+39 {numero_telefono.replace(' ', '')}"

        # Dati finali
        row = [
            sAMAccountName, "SI", OU, sAMAccountName, DisplayName, cn, nome_cap, cognome_cap,
            CF, "", "Utente esterno", "<PC>", "No", ExpireDate,
            userprincipalname, userprincipalname, mobile, "", "consip_vpn", "", "", "", ""
        ]

        # Creazione del nome file nel formato "<cognome>_<prima iniziale nome>.csv"
        nome_file = f"{cognome_cap}_{nome_cap[0]}.csv"

        # Scrittura su CSV
        with open(nome_file, "w", newline="") as file:
            writer = csv.writer(file, quoting=csv.QUOTE_MINIMAL)
            writer.writerow([
                "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
                "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
                "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
                "disable", "moveToOU", "telephoneNumber", "company"
            ])
            writer.writerow(row)

        st.success(f"File CSV generato correttamente come '{nome_file}' con data di scadenza '{ExpireDate}'")
        st.download_button(label="Scarica CSV", data=open(nome_file, "rb"), file_name=nome_file, mime="text/csv")
