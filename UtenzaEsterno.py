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

    # ---- BLOCCO AZURE ----
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

    # ---- BLOCCO DIPENDENTI e ESTERNI ----
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

        if tipo_utente == "Dipendente Consip":
            ou = st.selectbox("OU", ["Utenti standard", "Utenti VIP"], key="OU")
            employee_id = st.text_input("Employee ID", "", key="Employee ID").strip()
            department = st.text_input("Dipartimento", "", key="Dipartimento").strip()
            inserimento_gruppo = (
                "consip_vpn;dipendenti_wifi;mobile_wifi;"
                "GEDOGA-P-DOCGAR;GRPFreeDeskUser"
            )
            telephone_number = "+39 06 854491"
            company = "Consip"
        else:
            dip = st.selectbox(
                "Tipo di Esterno:",
                ["Consulente", "Somministrato/Stage"],
                key="tipo_esterno"
            )
            expire_date = st.text_input("Data di Fine (gg-mm-aaaa)", "30-06-2025", key="Data di Fine").strip()
            ou = (
                "Utenti esterni - Consulenti"
                if dip == "Consulente" else
                "Utenti esterni - Somministrati e Stage"
            )
            employee_id = ""
            if dip == "Somministrato/Stage":
                department = st.text_input("Dipartimento", "", key="Dipartimento").strip()
                inserimento_gruppo = (
                    "consip_vpn;dipendenti_wifi;mobile_wifi;GRPFreeDeskUser"
                )
            else:
                department = "Utente esterno"
                inserimento_gruppo = "consip_vpn"
            telephone_number = ""
            company = ""
            email_flag = (
                st.radio("Email necessaria?", ["Sì", "No"], key="flag_email") == "Sì"
            )
            if dip == "Consulente" and email_flag:
                try:
                    email = f"{cognome.lower()}{nome[0].lower()}@consip.it"
                except IndexError:
                    st.error("Per email automatica inserisci Nome e Cognome.")
                    email = ""
            elif dip == "Consulente":
                email = st.text_input("Email Personalizzata", "", key="Email").strip()
                inserimento_gruppo = "O365 Office App"
            else:
                email = (
                    f"{genera_samaccountname(
                        nome,
                        cognome,
                        secondo_nome,
                        segundo_cognome
                    )}@consip.it"
                )

        if st.button("Genera CSV"): 
            esterno = (tipo_utente == "Esterno")
            sAMAccountName = genera_samaccountname(
                nome, cognome, secondo_nome, secondo_cognome, esterno
            )
            nome_completo = ' '.join(
                [nome, secondo_nome, cognome, secondo_cognome]
            ).strip()
            display_name = (
                f"{nome_completo} (esterno)" if esterno else nome_completo
            )
            expire_fmt = formatta_data(expire_date) if esterno else ""
            upn = f"{sAMAccountName}@consip.it"
            mobile = f"+39 {numero_telefono}" if numero_telefono else ""
            description = description_input or "<PC>"
            mail = (
                email if (
                    tipo_utente == "Esterno" and dip == "Consulente" and not email_flag
                ) else f"{sAMAccountName}@consip.it"
            )
            given = ' '.join([nome, secondo_nome]).strip()
            surn = ' '.join([cognome, secondo_cognome]).strip()

            row = [
                sAMAccountName, "SI", ou, nome_completo, display_name,
                nome_completo, given, surn, codice_fiscale,
                employee_id, department, description, "No", expire_fmt,
                upn, mail, mobile, "", inserimento_gruppo, "", "",
                telephone_number, company
            ]

            buf = io.StringIO()
            wr = csv.writer(buf, quoting=csv.QUOTE_MINIMAL)
            wr.writerow(header_modifica)
            wr.writerow(row)
            buf.seek(0)

            df = pd.DataFrame([row], columns=header_modifica)
            st.dataframe(df)

            st.download_button(
                label="📥 Scarica CSV Utente",
                data=buf.getvalue(),
                file_name=f"{cognome}_{nome[:1]}_utente.csv",
                mime="text/csv"
            )
            st.success(f"✅ File CSV generato per '{sAMAccountName}'")

elif funzionalita == "Gestione Modifiche AD":
    st.subheader("Gestione Modifiche AD")
    # Campo per il nome del file di download
    file_name_modifiche = st.text_input(
        "Come si deve chiamare il file?",
        "modifiche_utenti.csv",
        key="FileName_Modifiche"
    )

    num_righe = st.number_input(
        "Quante righe vuoi inserire?", 1, 20, 1
    )
    modifiche = []
    for i in range(num_righe):
        with st.expander(f"Riga {i+1}"):
            m = {k: "" for k in header_modifica}
            m["sAMAccountName"] = st.text_input(
                f"[{i+1}] sAMAccountName *", key=f"user_{i}"
            )
            scelti = st.multiselect(
                f"[{i+1}] Campi da modificare", 
                [k for k in header_modifica if k != "sAMAccountName"],
                key=f"campi_{i}"
            )
            for campo in scelti:
                v = st.text_input(
                    f"[{i+1}] {campo}", key=f"{campo}_{i}"
                )
                if campo == "ExpireDate":
                    v = formatta_data(v)
                if campo == "mobile" and v and not v.startswith("+"):
                    v = f"+39 {v.strip()}"
                m[campo] = v
            modifiche.append(m)

    if st.button("Genera CSV Modifiche"):
        out = io.StringIO()
        w = csv.DictWriter(
            out, fieldnames=header_modifica, extrasaction='ignore'
        )
        w.writeheader()
        w.writerows(modifiche)
        out.seek(0)
        df2 = pd.DataFrame(modifiche, columns=header_modifica)
        st.dataframe(df2)

        # Messaggio di preview prima del download
        st.markdown(f"""
Ciao.

Si richiede modifica come da file {file_name_modifiche}
archiviato al percorso \\\\srv_dati.consip.tesoro.it\\AreaCondivisa\\DEPSI\\IC\\AD_Modifiche

Grazie
""")

        st.download_button(
            "📥 Scarica CSV Modifiche", 
            out.getvalue(), 
            file_name_modifiche, 
            "text/csv"
        )
        st.success("✅ CSV modifiche generato con successo.")
