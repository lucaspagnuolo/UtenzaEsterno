import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# Inizializza lo stato della sessione
reset_keys = [
    "Nome", "Secondo Nome", "Cognome", "Secondo Cognome", "Numero di Telefono", "Description", "Codice Fiscale",
    "Data di Fine", "Employee ID", "Dipartimento", "Email", "flag_email",
    # chiavi Azure
    "Nome_Azure", "SecondoNome_Azure", "Cognome_Azure", "SecondoCognome_Azure",
    "TelAziendale", "EmailAziendale", "Manager_Azure", "SM_Azure"
]

if "reset_fields" not in st.session_state:
    st.session_state.reset_fields = False

# Pulsante per pulire i campi
if st.button("🔄 Pulisci Campi"):
    st.session_state.reset_fields = True

# Reset campi
if st.session_state.reset_fields:
    for key in reset_keys:
        st.session_state[key] = ""
    st.session_state.reset_fields = False

# Funzioni di utilità
def formatta_data(data):
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(
    nome: str,
    cognome: str,
    secondo_nome: str = "",
    secondo_cognome: str = "",
    esterno: bool = False
) -> str:
    n = nome.strip().lower()
    sn = secondo_nome.strip().lower() if secondo_nome else ""
    c = cognome.strip().lower()
    sc = secondo_cognome.strip().lower() if secondo_cognome else ""
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{(n[0] if n else '')}{(sn[0] if sn else '')}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{(n[0] if n else '')}{(sn[0] if sn else '')}.{c}"
    return cand[:limit] + suffix

# Interfaccia
st.title("Gestione Utenze Consip")
funzionalita = st.radio(
    "Scegli funzionalità:",
    ["Gestione Creazione Utenze", "Gestione Modifiche AD"]
)

header_modifica = [
    "sAMAccountName", "Creation", "OU", "Name", "DisplayName", "cn", "GivenName", "Surname",
    "employeeNumber", "employeeID", "department", "Description", "passwordNeverExpired",
    "ExpireDate", "userprincipalname", "mail", "mobile", "RimozioneGruppo", "InserimentoGruppo",
    "disable", "moveToOU", "telephoneNumber", "company"
]

if funzionalita == "Gestione Creazione Utenze":
    tipo_utente = st.selectbox(
        "Seleziona il tipo di utente:",
        ["Dipendente Consip", "Esterno", "Azure"],
        key="tipo_utente"
    )

    if tipo_utente == "Azure":
        nome = st.text_input("Nome", key="Nome_Azure").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome", key="SecondoNome_Azure").strip().capitalize()
        cognome = st.text_input("Cognome", key="Cognome_Azure").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome", key="SecondoCognome_Azure").strip().capitalize()
        telefono_aziendale = st.text_input("Telefono Aziendale (senza prefisso)", key="TelAziendale").strip()
        email_aziendale = st.text_input("Email Aziendale", key="EmailAziendale").strip()
        manager = st.text_input("Manager", key="Manager_Azure").strip()
        sm_text = st.text_area("Sulle quali SM va profilato (uno per riga)", key="SM_Azure")
        sm_list = [s.strip() for s in sm_text.split("\n") if s.strip()]

        if st.button("Genera Richiesta Azure"):
            sAMAccountName = genera_samaccountname(
                nome,
                cognome,
                secondo_nome,
                secondo_cognome,
                esterno=True
            )
            telefono_fmt = f"+39 {telefono_aziendale}" if telefono_aziendale else ""
            # Messaggio iniziale
            st.markdown("Ciao.\nRichiedo cortesemente la definizione di una utenza su Azure come di sotto indicato.")
            # Calcolo Display Name con cognome+secondo_cognome, nome+secondo_nome, senza spazi vuoti multipli
            display_parts = [cognome + secondo_cognome, nome, secondo_nome]
            display_name_str = " ".join([part for part in display_parts if part]).strip() + " (esterno)"
            # Costruzione tabella
            table = [
                ["Campo", "Valore"],
                ["Tipo Utenza", "Azure"],
                ["Utenza", sAMAccountName],
                ["Alias", sAMAccountName],
                ["Nome", " ".join([nome, secondo_nome]).strip()],
                ["Cognome", " ".join([cognome, secondo_cognome]).strip()],
                ["Display name", display_name_str],
                ["Email aziendale", email_aziendale],
                ["Manager", manager],
                ["Cell", telefono_fmt],
                ["e-mail Consip", f"{sAMAccountName}@consip.it"]
            ]
            # Render tabella Markdown
            table_md = "| " + " | ".join(table[0]) + " |\n"
            table_md += "| " + " | ".join(["---"]*len(table[0])) + " |\n"
            for row in table[1:]:
                table_md += "| " + " | ".join(row) + " |\n"
            st.markdown(table_md)

            # Dettagli aggiuntivi
            st.markdown("""
Aggiungere all’utenza la MFA

Aggiungere all’utenza le licenze:
- Microsoft Defender per Office 365 (piano 2)
- Office 365 E3
""")

            if sm_list:
                st.markdown("Profilare su SM:")
                for sm in sm_list:
                    st.markdown(f"- {sm}@consip.it")

            # Invio credenziali
            st.markdown(f"""
La comunicazione delle credenziali dovranno essere inviate:
- utenza via email a {email_aziendale}
- psw via SMS a {telefono_fmt}
""")
            # URL web mail
            if sm_list:
                for sm in sm_list:
                    st.markdown(f"La url per la web mail è https://outlook.office.com/mail/{sm}@consip.it")
            st.markdown("Grazie")

    else:
        nome = st.text_input("Nome", key="Nome").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome", key="Secondo Nome").strip().capitalize()
        cognome = st.text_input("Cognome", key="Cognome").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome", key="Secondo Cognome").strip().capitalize()
        numero_telefono = st.text_input("Numero di Telefono", "", key="Numero di Telefono").replace(" ", "")
        description_input = st.text_input(
            "Description (lascia vuoto per <PC>)",
            "<PC>",
            key="Description"
        ).strip()
        codice_fiscale = st.text_input("Codice Fiscale", "", key="Codice Fiscale").strip()

        # (resto del codice inalterato...)
