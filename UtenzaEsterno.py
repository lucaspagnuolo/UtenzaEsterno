import streamlit as st
import pandas as pd
import io
import csv
from datetime import datetime

st.set_page_config(page_title="Generatore CSV Active Directory", layout="wide")
st.title("🔐 Generatore CSV Active Directory")

header_modifica = [
    "sAMAccountName", "creaUtente", "OU", "CN", "displayName", "name", "givenName", "sn",
    "employeeNumber", "employeeID", "department", "description", "changePasswordAtLogon",
    "ExpireDate", "userPrincipalName", "mail", "mobile", "manager", "Gruppi", "Gruppi da rimuovere",
    "Gruppi da aggiungere", "telephoneNumber", "company"
]

def formatta_data(data_str):
    try:
        return datetime.strptime(data_str.strip(), "%d/%m/%Y").strftime("%Y%m%d")
    except:
        return data_str

def genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno=False):
    nome = nome.lower().replace(" ", "")
    cognome = cognome.lower().replace(" ", "")
    secondo_nome = secondo_nome.lower().replace(" ", "")
    secondo_cognome = secondo_cognome.lower().replace(" ", "")
    base = f"{nome[0]}{cognome}"
    if secondo_nome:
        base += secondo_nome[0]
    if secondo_cognome:
        base += secondo_cognome[0]
    if esterno:
        base += "_ext"
    return base

funzionalita = st.sidebar.selectbox("Scegli Funzionalità", ["Creazione Utente", "Gestione Modifiche AD"])

if funzionalita == "Creazione Utente":
    tipo_utente = st.radio("Tipo di Utenza", ["Dipendente Consip", "Esterno", "Azure"])

    if tipo_utente == "Dipendente Consip":
        nome = st.text_input("Nome", key="Nome").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome", key="Secondo Nome").strip().capitalize()
        cognome = st.text_input("Cognome", key="Cognome").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome", key="Secondo Cognome").strip().capitalize()
        codice_fiscale = st.text_input("Codice Fiscale", "", key="Codice Fiscale").strip()
        ou = "Utenti Consip"
        employee_id = st.text_input("Employee ID", "", key="Employee ID").strip()
        department = st.text_input("Dipartimento", "", key="Dipartimento").strip()
        description = "<CD>"
        inserimento_gruppo = "consip_wifi"
        telephone_number = "+39 06 854491"
        company = "Consip"
        email_flag = st.checkbox("Vuoi ricevere un indirizzo email?", key="flag_email")
        email = st.text_input("Email", key="Email").strip()
        employee_number = codice_fiscale

        if st.button("Genera CSV"):
            esterno = False
            sAMAccountName = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno)

            nome_completo = f"{nome} {secondo_nome} {cognome} {secondo_cognome}".strip()
            nome_completo = ' '.join(nome_completo.split())

            display_name = f"{nome_completo}"
            userprincipalname = f"{sAMAccountName}@consip.it" if email_flag else ""
            mail = f"{sAMAccountName}@consip.it" if email_flag else ""

            given_name = f"{nome} {secondo_nome}".strip()
            surname = f"{cognome} {secondo_cognome}".strip()

            row = [
                sAMAccountName, "SI", ou, nome_completo, display_name, nome_completo, given_name, surname,
                employee_number, employee_id, department, description, "No", "",
                userprincipalname, mail, "", "", inserimento_gruppo, "", "", telephone_number, company
            ]

            output_main = io.StringIO()
            writer_main = csv.writer(output_main)
            writer_main.writerow(header_modifica)
            writer_main.writerow(row)

            st.download_button(
                label="📥 Scarica CSV",
                data=output_main.getvalue(),
                file_name=f"{sAMAccountName}_creazione.csv",
                mime="text/csv"
            )
    elif tipo_utente == "Esterno":
        nome = st.text_input("Nome").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
        cognome = st.text_input("Cognome").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
        societa = st.text_input("Società Esterna").strip()
        ou = "Utenti Esterni"
        manager = st.text_input("Manager").strip()
        mobile = st.text_input("Numero Mobile").strip()
        email = st.text_input("Email Aziendale").strip()
        description = societa
        telephone_number = ""
        company = societa

        if st.button("Genera CSV Esterno"):
            esterno = True
            sAMAccountName = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno)

            nome_completo = f"{nome} {secondo_nome} {cognome} {secondo_cognome}".strip()
            nome_completo = ' '.join(nome_completo.split())

            display_name = f"{nome_completo} (esterno)"
            userprincipalname = f"{sAMAccountName}@consip.it"
            mail = email

            given_name = f"{nome} {secondo_nome}".strip()
            surname = f"{cognome} {secondo_cognome}".strip()

            row = [
                sAMAccountName, "SI", ou, nome_completo, display_name, nome_completo, given_name, surname,
                "", "", "", description, "No", "",
                userprincipalname, mail, mobile, manager, "", "", "", telephone_number, company
            ]

            output_ext = io.StringIO()
            writer_ext = csv.writer(output_ext)
            writer_ext.writerow(header_modifica)
            writer_ext.writerow(row)

            st.download_button(
                label="📥 Scarica CSV Esterno",
                data=output_ext.getvalue(),
                file_name=f"{sAMAccountName}_esterno.csv",
                mime="text/csv"
            )

    elif tipo_utente == "Azure":
        nome = st.text_input("Nome").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
        cognome = st.text_input("Cognome").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
        email_aziendale = st.text_input("Email Aziendale").strip()
        manager = st.text_input("Manager").strip()
        mobile = st.text_input("Mobile").strip()
        profilazioni_sm = st.text_area("Profilare sulla SM (una per riga)").strip().splitlines()

        if st.button("Genera Testo Richiesta Azure"):
            sAMAccountName = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno=False)
            nome_full = f"{nome} {secondo_nome}".strip()
            cognome_full = f"{cognome} {secondo_cognome}".strip()
            display_name = f"{cognome} {nome} (esterno)"
            mail_consip = f"{sAMAccountName}@consip.it"

            testo = f"""Ciao.
Richiedo cortesemente la definizione di una utenza su azure come di sotto indicato.

Tipo Utenza\tAzure
Utenza\t{sAMAccountName}
Alias\t{sAMAccountName}
Nome\t{nome} {secondo_nome}
Cognome\t{cognome} {secondo_cognome}
Display name\t{display_name}
Email aziendale\t{email_aziendale}
Manager\t{manager}
Cell\t{mobile}
e-mail Consip\t{mail_consip}

Aggiungere all’utenza la MFA
Aggiungere all’utenza le licenze:
•\tMicrosoft Defender per Office 365 (piano 2)
•\tOffice 365 E3"""

            for profilo in profilazioni_sm:
                if profilo.strip():
                    testo += f"\nProfilare su SM {profilo.strip()}"

            testo += f"""

La comunicazioni delle credenziali dovranno essere inviate:
•\tutenza via email a {email_aziendale}
•\tpsw via SMS a {mobile}
la url per la web mail è https://outlook.office.com/mail/{email_aziendale}
Grazie
"""
            st.text_area("📄 Testo Richiesta Utenza Azure", testo, height=500)
else:
    st.subheader("Gestione Modifiche AD")

    num_righe = st.number_input("Quante righe vuoi inserire?", min_value=1, max_value=20, value=1, step=1)
    modifiche = []

    for i in range(num_righe):
        with st.expander(f"Riga {i+1}"):
            modifica = {key: "" for key in header_modifica}
            modifica["sAMAccountName"] = st.text_input(f"[{i+1}] sAMAccountName *", key=f"user_{i}")

            campi_selezionati = st.multiselect(
                f"[{i+1}] Seleziona i campi da modificare",
                [k for k in header_modifica if k != "sAMAccountName"],
                key=f"campi_{i}"
            )

            for campo in campi_selezionati:
                valore = st.text_input(f"[{i+1}] {campo}", key=f"{campo}_{i}")
                if campo == "ExpireDate":
                    valore = formatta_data(valore)
                elif campo == "mobile":
                    valore = valore.strip()
                    if valore and not valore.startswith("+"):
                        valore = f"+39 {valore}"
                modifica[campo] = valore

            modifiche.append(modifica)

    if st.button("Genera CSV Modifiche"):
        output_modifiche = io.StringIO()
        writer = csv.DictWriter(output_modifiche, fieldnames=header_modifica, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(modifiche)
        output_modifiche.seek(0)

        df_modifiche = pd.DataFrame(modifiche, columns=header_modifica)
        st.dataframe(df_modifiche)

        st.download_button(
            label="📥 Scarica CSV Modifiche",
            data=output_modifiche.getvalue(),
            file_name="modifiche_utenti.csv",
            mime="text/csv"
        )

        st.success("✅ CSV modifiche generato con successo.")
