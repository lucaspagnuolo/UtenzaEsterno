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
    "TelAziendale", "EmailAziendale", "Manager_Azure", "SM_Azure", "Data_fine_Azure"
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


def build_full_name(
    cognome: str,
    secondo_cognome: str,
    nome: str,
    secondo_nome: str,
    esterno: bool = False
) -> str:
    parts = [cognome, secondo_cognome, nome, secondo_nome]
    parts = [p for p in parts if p]
    full = ' '.join(parts)
    if esterno:
        full += ' (esterno)'
    return full

# Interfaccia
st.title("Gestione Utenze Consip")
funzionalita = st.radio(
    "Scegli funzionalità:",
    ["Gestione Creazione Utenze", "Gestione Modifiche AD", "Deprovisioning"]
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

    # Azure
    if tipo_utente == "Azure":
        nome = st.text_input("Nome", key="Nome_Azure").strip().capitalize()
        secondo_nome = st.text_input("Secondo Nome", key="SecondoNome_Azure").strip().capitalize()
        cognome = st.text_input("Cognome", key="Cognome_Azure").strip().capitalize()
        secondo_cognome = st.text_input("Secondo Cognome", key="SecondoCognome_Azure").strip().capitalize()
        telefono_aziendale = st.text_input("Telefono Aziendale (senza prefisso)", key="TelAziendale").strip()
        email_aziendale = st.text_input("Email Aziendale", key="EmailAziendale").strip()
        manager = st.text_input("Manager", key="Manager_Azure").strip()
        data_fine = st.text_input("Data Fine (gg/mm/aaaa)", key="Data_fine_Azure").strip()        

        casella_personale = st.checkbox("Casella Personale Consip", key="Casella_Personale_Azure")
        sm_list = []
        if casella_personale:
            sm_text = st.text_area("Sulle quali SM va profilato (uno per riga)", key="SM_Azure")
            sm_list = [s.strip() for s in sm_text.split("\n") if s.strip()]

        if st.button("Genera Richiesta Azure"):
            sAMAccountName = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, esterno=True)
            telefono_fmt = f"+39 {telefono_aziendale}" if telefono_aziendale else ""
            display_name_str = build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=True)

            # Costruzione tabella base
            table = [
                ["Campo", "Valore"],
                ["Tipo Utenza", "Azure"],
                ["Utenza", sAMAccountName],
                ["Alias", sAMAccountName],
                ["Name", build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False)],
                ["DisplayName", display_name_str],
                ["cn", display_name_str],
                ["GivenName", ' '.join([nome, secondo_nome]).strip()],
                ["Surname", ' '.join([cognome, secondo_cognome]).strip()],
                ["Email aziendale", email_aziendale],
                ["Manager", manager],
                ["Cell", telefono_fmt],
                ["Data Fine (mm/gg/aaaa)",formatta_data(data_fine)]
            ]
            # Campo extra se Casella Personale
            if casella_personale:
                table.append(["e-mail Consip", f"{sAMAccountName}@consip.it"])

            # Render Markdown
            table_md = "| " + " | ".join(table[0]) + " |\n"
            table_md += "| " + " | ".join(["---"]*len(table[0])) + " |\n"
            for row in table[1:]:
                table_md += "| " + " | ".join(row) + " |\n"
            st.markdown("""Ciao, Si richiede la definizione di un utenza Azure come sotto indicato.""")
            st.markdown(table_md)

            # Licenze e SM se Casella Personale
            if casella_personale:
                st.markdown("""
Aggiungere all’utenza le licenze:
- Microsoft Defender per Office 365 (piano 2)
- Office 365 E3
""")
                if sm_list:
                    st.markdown("Profilare su SM:")
                    for sm in sm_list:
                        st.markdown(f"- {sm}@consip.it")

            # MFA e credenziali
            st.markdown("""
Aggiungere all’utenza la MFA

Gli utenti verranno contattati e supportati per accesso e configurazione MFA da imac@consip.it:
- utenza via email a {email_aziendale}
- psw via SMS a {telefono_fmt}

Si richiede cortesemente contatto utenti per MFA/accesso utente/webmail:
{display_name_str}-   {telefono_fmt} -  {email_aziendale}. Nota attenzione alla PSW. Grazie ciao
""".format(email_aziendale=email_aziendale, telefono_fmt=telefono_fmt,display_name_str=display_name_str))
            if casella_personale and sm_list:
                for sm in sm_list:
                    st.markdown(f"La url per la web mail è https://outlook.office.com/mail/{sm}@consip.it")
            st.markdown("Grazie")
    else:
        # ---- BLOCCO DIPENDENTI E ESTERNI ----
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
                        secondo_cognome
                    )}@consip.it"
                )

        if st.button("Genera CSV"): 
            esterno = (tipo_utente == "Esterno")
            sAMAccountName = genera_samaccountname(
                nome, cognome, secondo_nome, secondo_cognome, esterno
            )
            cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno)
            nome_completo = cn.replace(" (esterno)", "")
            display_name = cn
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
                cn, given, surn, codice_fiscale,
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

        file_path = "\\srv_dati.consip.tesoro.it\\AreaCondivisa\\DEPSI\\IC\\AD_Modifiche"
        st.markdown(
            f"""Ciao.

Si richiede modifica come da file {file_name_modifiche}
archiviato al percorso {file_path}

Grazie"""
        )

        st.download_button(
            "📥 Scarica CSV Modifiche", 
            out.getvalue(), 
            file_name_modifiche, 
            "text/csv"
        )
        st.success("✅ CSV modifiche generato con successo.")
elif funzionalita == "Deprovisioning":
    st.subheader("Deprovisioning Utente")

    sam = st.text_input("Nome utente (sAMAccountName)", "").strip().lower()
    st.markdown("---")

    dl_file = st.file_uploader("Carica file DL (Excel)", type="xlsx")
    sm_file = st.file_uploader("Carica file SM (Excel)", type="xlsx")
    mg_file = st.file_uploader("Carica file Membri_Gruppi (Excel)", type="xlsx")

    if st.button("Genera Deprovisioning"):
        dl_df = pd.read_excel(dl_file) if dl_file else pd.DataFrame()
        sm_df = pd.read_excel(sm_file) if sm_file else pd.DataFrame()
        mg_df = pd.read_excel(mg_file) if mg_file else pd.DataFrame()

        dl_list = []
        if not dl_df.empty and dl_df.shape[1] > 5:
            mask = dl_df.iloc[:, 1].astype(str).str.lower() == sam
            dl_list = dl_df.loc[mask, dl_df.columns[5]].dropna().tolist()

        sm_list = []
        if not sm_df.empty and sm_df.shape[1] > 2:  # Controllo che ci siano almeno 3 colonne
            target = f"{sam}@consip.it"
            mask = sm_df.iloc[:, 2].astype(str).str.lower() == target  # filtro sulla terza colonna
            sm_list = sm_df.loc[mask, sm_df.columns[0]].dropna().tolist()  # estrazione dalla prima colonna


        grp = []
        if not mg_df.empty and mg_df.shape[1] > 3:
            mask = mg_df.iloc[:, 3].astype(str).str.lower() == sam
            grp = mg_df.loc[mask, mg_df.columns[0]].dropna().tolist()

        lines = [
            f"Ciao,\nper {sam}@consip.it :"
        ]
        warnings = []
        step = 1

        lines.append(f"{step}. Disabilitare invio ad utente (Message Delivery Restrictions)")
        step += 1
        lines.append(f"{step}. Impostare Hide dalla Rubrica")
        step += 1
        lines.append(f"{step}. Disabilitare accesso Mailbox (Mailbox features – Disable Protocolli/OWA)")
        step += 1
        lines.append(f"{step}. Estrarre il PST (O365 eDiscovery) da archiviare in \\\\nasconsip2....\\backuppst\\03 - backup email cancellate\\{sam}@consip.it (in z7 con psw condivisa)")
        step += 1
        lines.append(f"{step}. Rimuovere le appartenenze dall’utenza Azure")
        step += 1
        lines.append(f"{step}. Rimuovere le applicazioni dall’utenza Azure")
        step += 1

        if dl_list:
            lines.append(f"{step}. Rimozione abilitazione dalle DL")
            for dl in dl_list:
                lines.append(f"   - {dl}")
            step += 1
        else:
            warnings.append("⚠️ Non sono state trovate DL all'utente indicato")

        lines.append(f"{step}. Disabilitare l’account di Azure")
        step += 1

        if sm_list:
            lines.append(f"{step}. Rimozione abilitazione da SM")
            for sm in sm_list:
                lines.append(f"   - {sm}")
            step += 1
        else:
            warnings.append("⚠️ Non sono state trovate SM profilate all'utente indicato")

        lines.append(f"{step}. Rimozione in AD del gruppo")
        lines.append("   - O365 Copilot Plus")
        lines.append("   - O365 Teams Premium")
        utenti_groups = [g for g in grp if g.lower().startswith("o365 utenti")]
        if utenti_groups:
            for g in utenti_groups:
                lines.append(f"   - {g}")
        else:
            warnings.append("⚠️ Non è stato trovato nessun gruppo O365 Utenti per l'utente")
        step += 1

        lines.append(f"{step}. Disabilitazione utenza di dominio")
        step += 1
        lines.append(f"{step}. Spostamento in dismessi/utenti")
        step += 1
        lines.append(f"{step}. Cancellare la foto da Azure (se applicabile)")
        step += 1
        lines.append(f"{step}. Rimozione Wi-Fi")

        # Aggiunta dei warning alla fine, se presenti
        if warnings:
            lines.append("")  # Riga vuota
            lines.append("⚠️ Avvisi:")
            lines.extend(warnings)

        st.text("\n".join(lines))






