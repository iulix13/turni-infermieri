import streamlit as st
import pandas as pd
import calendar
import io

st.set_page_config(page_title="Turnistica Mensile Infermieri - Aziendale", layout="wide")

MESI_IT = {
    1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE",
    5: "MAGGIO", 6: "GIUGNO", 7: "LUGLIO", 8: "AGOSTO",
    9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
}

st.title("📋 Pianificatore Turni Infermieristici (Modello Aziendale & Target 156h)")

# --- SIDEBAR: CONFIGURAZIONE MESE E REPARTO ---
st.sidebar.header("1. Impostazioni Mese e Reparto")
intestazione_dip = st.sidebar.text_input("Intestazione Dipartimento:", value="DIREZIONE INFERMIERISTICA E TECNICA DIPARTIMENTO CHIRURGICO")
reparto = st.sidebar.text_input("Nome Reparto:", value="CHIRURGIA DEGENZE")
anno = st.sidebar.number_input("Anno", min_value=2024, max_value=2030, value=2026)
mese = st.sidebar.selectbox("Mese", list(range(1, 13)), index=10, format_func=lambda x: MESI_IT[x])
data_emissione = st.sidebar.text_input("Data Emissione:", value="20/01/2026")

giorni_nel_mese = calendar.monthrange(anno, mese)[1]
nomi_giorni_it = ["L", "M", "M", "G", "V", "S", "D"]

st.sidebar.header("2. Fabbisogno Giornaliero Reparto")
req_m = st.sidebar.number_input("Infermieri Mattina (M):", min_value=0, max_value=50, value=6)
req_p = st.sidebar.number_input("Infermieri Pomeriggio (P):", min_value=0, max_value=50, value=6)
req_n = st.sidebar.number_input("Infermieri Notte (N):", min_value=0, max_value=50, value=4)

# --- SIDEBAR: CONFIGURAZIONE PERSONALE ---
st.sidebar.header("3. Configurazione Personale Reale")
num_infermieri = st.sidebar.number_input("Numero totale infermieri in servizio:", min_value=1, max_value=100, value=23)

modalita_personale = st.sidebar.radio(
    "Inserimento Nomi:",
    ["Inserisci Nomi Reali", "Carica File Excel/CSV"]
)

# Inizializzazione Session State per il Personale (Default 23 Unità con Elisa e Domenico)
if 'df_pers' not in st.session_state or len(st.session_state.df_pers) != num_infermieri:
    nomi_default = [
        "ELISA", "DOMENICO", "MARIO", "ANNA", "GIUSEPPE", 
        "FRANCESCA", "LUCA", "SARA", "ALESSANDRO", "MARTINA", 
        "ROBERTO", "ELENA", "GIOVANNI", "CHIARA", "STEFANO", 
        "FEDERICA", "ANTONIO", "VALENTINA", "PAOLO", "GIULIA", 
        "MARCO", "SILVIA", "FEDERICO"
    ]
    
    dati_iniziali = []
    for idx in range(num_infermieri):
        nome = nomi_default[idx] if idx < len(nomi_default) else f"Infermiere {idx+1}"
        
        # Impostazioni per Elisa, Domenico e personale speciale
        if nome in ["ELISA", "DOMENICO"]:
            es_notte = "Sì"
            t_fisso = "Solo Mattina (M)"
        elif idx == 2:
            es_notte = "Sì"
            t_fisso = "Nessuno"
        else:
            es_notte = "No"
            t_fisso = "Nessuno"
            
        dati_iniziali.append({
            "Nome": nome,
            "Qualifica": "Infermiere",
            "Esonero_Notte": es_notte,
            "Turno_Fisso": t_fisso,
            "Turno_Inizio_Mese": "M"
        })
    st.session_state.df_pers = pd.DataFrame(dati_iniziali)

if modalita_personale == "Inserisci Nomi Reali":
    st.sidebar.markdown("#### Nomi dell'Unità Operativa")
    nomi_correnti = "\n".join(st.session_state.df_pers["Nome"].tolist())
    testo_nomi = st.sidebar.text_area("Inserisci i nomi (uno per riga):", value=nomi_correnti, height=200)
    lista_nomi_inseriti = [n.strip() for n in testo_nomi.split("\n") if n.strip()]
    
    if len(lista_nomi_inseriti) == num_infermieri:
        st.session_state.df_pers["Nome"] = lista_nomi_inseriti
    elif len(lista_nomi_inseriti) > 0 and len(lista_nomi_inseriti) != num_infermieri:
        st.sidebar.warning(f"Hai inserito {len(lista_nomi_inseriti)} nomi ma l'organico è di {num_infermieri}.")
else:
    file_personale = st.sidebar.file_uploader("Carica file Excel/CSV Personale", type=["xlsx", "csv"])
    if file_personale is not None:
        try:
            if file_personale.name.endswith('.csv'):
                df_temp = pd.read_csv(file_personale)
            else:
                df_temp = pd.read_excel(file_personale)
            for col in ["Esonero_Notte", "Turno_Fisso", "Turno_Inizio_Mese"]:
                if col not in df_temp.columns:
                    df_temp[col] = "No" if col == "Esonero_Notte" else ("Nessuno" if col == "Turno_Fisso" else "M")
            st.session_state.df_pers = df_temp
            st.sidebar.success("File personale caricato!")
        except Exception as e:
            st.sidebar.error(f"Errore file: {e}")

# --- SCHEDE PRINCIPALI ---
tab_personale, tab_richieste, tab_matrice = st.tabs([
    "👥 1. Anagrafica, Esoneri & Turni Fissi", 
    "📌 2. Gestione Richieste, Ferie & Desiderata", 
    "📊 3. Matrice Turni Validata"
])

# --- TAB 1: GESTIONE PERSONALE ---
with tab_personale:
    st.subheader(f"Gestione Personale in Servizio ({len(st.session_state.df_pers)} Infermieri)")
    st.info("💡 Imposta esoneri notturni, turni fissi (es. Solo Mattina per Elisa e Domenico) e turno iniziale. Clicca su **Salva Modifiche**.")
    
    with st.form("form_personale"):
        edited_df = st.data_editor(
            st.session_state.df_pers,
            num_rows="dynamic",
            column_config={
                "Nome": st.column_config.TextColumn("Nome Infermiere", required=True),
                "Qualifica": st.column_config.TextColumn("Qualifica"),
                "Esonero_Notte": st.column_config.SelectboxColumn(
                    "Esonero Notte",
                    options=["Sì", "No"],
                    required=True
                ),
                "Turno_Fisso": st.column_config.SelectboxColumn(
                    "Turno Fisso (Esclusivo)",
                    options=["Nessuno", "Solo Mattina (M)", "Solo Pomeriggio (P)"],
                    required=True
                ),
                "Turno_Inizio_Mese": st.column_config.SelectboxColumn(
                    "Turno 1° del Mese",
                    options=["M", "P", "N", "S", "R", "F", "PR", "104", "CFD", "PC"],
                    required=True
                )
            },
            use_container_width=True,
            height=480
        )
        
        salva_btn = st.form_submit_button("💾 Salva Modifiche Personale", type="primary")
        if salva_btn:
            st.session_state.df_pers = edited_df
            st.success("Modifiche salvate con successo!")

# --- TAB 2: RICHIESTE MESE, FERIE & DESIDERATA ---
with tab_richieste:
    st.subheader("Inserimento Richieste, Assenze e Desiderata")
    
    if 'eccezioni' not in st.session_state:
        st.session_state.eccezioni = {}  
        
    list_nomi = st.session_state.df_pers["Nome"].tolist() if not st.session_state.df_pers.empty else []
    
    col_r1, col_r2, col_r3, col_r4 = st.columns([2, 2, 1.5, 1.5])
    with col_r1:
        inf_scelto = st.selectbox("Seleziona Infermiere:", list_nomi if list_nomi else ["Nessuno"])
    with col_r2:
        causale_scelta = st.selectbox(
            "Tipo Evento / Desiderato:", 
            [
                "F (Ferie)", "PR (Permesso)", "104 (Legge 104)", "FE (Festivo)",
                "ASP (Aspettativa)", "REC (Recupero Ore)", "CFD (Centro Ferite Difficili - P)", "PC (Prericovero Centralizzato - M)",
                "M_DES (Desiderata Mattina)", "P_DES (Desiderata Pomeriggio)", "N_DES (Desiderata Notte)", "R (Riposo forzato)"
            ]
        )
        sigla_causale = causale_scelta.split(" ")[0]
    with col_r3:
        giorno_inizio = st.number_input("Dal giorno:", min_value=1, max_value=giorni_nel_mese, value=1)
    with col_r4:
        giorno_fine = st.number_input("Al giorno:", min_value=1, max_value=giorni_nel_mese, value=1)

    if st.button("➕ Registra Richiesta", type="primary"):
        if giorno_inizio <= giorno_fine:
            for g in range(giorno_inizio, giorno_fine + 1):
                st.session_state.eccezioni[(inf_scelto, g)] = sigla_causale
            st.success(f"Registrato '{sigla_causale}' per {inf_scelto} (gg. {giorno_inizio}-{giorno_fine}).")
        else:
            st.error("Il giorno di inizio deve essere precedente o uguale al giorno di fine.")

    st.markdown("---")
    st.subheader("📋 Elenco Richieste Attive")
    
    if st.session_state.eccezioni:
        lista_ecc = []
        for (inf, g), causale in st.session_state.eccezioni.items():
            lista_ecc.append({"Infermiere": inf, "Giorno": g, "Causale": causale})
        
        df_ecc = pd.DataFrame(lista_ecc).sort_values(by=["Infermiere", "Giorno"])
        
        for idx, row in df_ecc.iterrows():
            c_inf, c_g, c_caus, c_act = st.columns([3, 2, 2, 1])
            c_inf.write(f"**{row['Infermiere']}**")
            c_g.write(f"Giorno {row['Giorno']}")
            c_caus.write(f"`{row['Causale']}`")
            if c_act.button("❌", key=f"del_{row['Infermiere']}_{row['Giorno']}"):
                del st.session_state.eccezioni[(row['Infermiere'], row['Giorno'])]
                st.rerun()
                
        st.write("")
        if st.button("🗑️ Cancella Tutte le Richieste"):
            st.session_state.eccezioni = {}
            st.rerun()
    else:
        st.info("Nessuna richiesta o desiderata presente al momento.")

# --- TAB 3: GENERATORE MATRICE AZIENDALE VALIDATA ---
with tab_matrice:
    st.markdown("""
        <style>
        .header-box { border: 2px solid #000; padding: 6px; background-color: #f9f9f9; text-align: center; font-weight: bold; }
        .header-title { font-size: 15px; text-transform: uppercase; }
        .sub-header { border-left: 2px solid #000; border-right: 2px solid #000; border-bottom: 2px solid #000; padding: 4px; font-weight: bold; background-color: #f0f0f0; }
        </style>
    """, unsafe_allow_html=True)

    # Intestazione formale in stile Matrice Ufficiale
    c1, c2, c3 = st.columns([3, 2, 2.5])
    with c1: st.markdown(f"<div class='header-box'>{intestazione_dip}</div>", unsafe_allow_html=True)
    with c2: st.markdown(f"<div class='header-box'>DATA EMISSIONE: {data_emissione}</div>", unsafe_allow_html=True)
    with c3: st.markdown("<div class='header-box header-title'>ORE DOVUTE: 156</div>", unsafe_allow_html=True)

    cs1, cs2 = st.columns([3, 2])
    with cs1: st.markdown(f"<div class='sub-header'>{reparto}</div>", unsafe_allow_html=True)
    with cs2: st.markdown(f"<div class='sub-header'>MESE DI {MESI_IT[mese]} {anno}</div>", unsafe_allow_html=True)

    st.write("")

    giorni_settimana = []
    colonne_domeniche = []
    for g in range(1, giorni_nel_mese + 1):
        giorno_sett_num = calendar.weekday(anno, mese, g)
        giorni_settimana.append(nomi_giorni_it[giorno_sett_num])
        if giorno_sett_num == 6:
            colonne_domeniche.append(str(g))

    r_num = {str(g): g for g in range(1, giorni_nel_mese + 1)}
    r_num["TOT ore"] = ""
    r_let = {str(g): giorni_settimana[g - 1] for g in range(1, giorni_nel_mese + 1)}
    r_let["TOT ore"] = ""
    st.dataframe(pd.DataFrame([r_num, r_let], index=["Giorno", "Settimana"]), use_container_width=True)

    df_pers_attivo = st.session_state.df_pers
    if df_pers_attivo is not None and not df_pers_attivo.empty:
        nomi_inf_list = [str(r["Nome"]).strip() for _, r in df_pers_attivo.iterrows()]
        
        matrice_turni = {nome: {} for nome in nomi_inf_list}
        esoneri_dict = {
            str(r["Nome"]).strip(): (str(r.get("Esonero_Notte", "No")).strip().lower() in ["sì", "si", "yes", "true", "1"])
            for _, r in df_pers_attivo.iterrows()
        }
        turni_fissi_dict = {
            str(r["Nome"]).strip(): str(r.get("Turno_Fisso", "Nessuno")).strip()
            for _, r in df_pers_attivo.iterrows()
        }
        turni_inizio_dict = {
            str(r["Nome"]).strip(): str(r.get("Turno_Inizio_Mese", "M")).strip()
            for _, r in df_pers_attivo.iterrows()
        }

        pesi_ore = {
            "M": 7, "P": 7, "S": 7, "CFD": 7, "PC": 7, "M_DES": 7, "P_DES": 7,
            "N": 3, "N_DES": 3,
            "F": 6, "PR": 6, "104": 6, "FE": 6, "ASP": 6, "REC": 6,
            "R": 0
        }

        # Matrice sfalsata dinamica (Ciclo M -> P -> N -> S -> R)
        rotazione = ["M", "P", "N", "S", "R"]
        
        for idx, nome in enumerate(nomi_inf_list):
            offset = (idx * 2) % len(rotazione)
            
            for g in range(1, giorni_nel_mese + 1):
                str_g = str(g)
                tf = turni_fissi_dict.get(nome, "Nessuno")
                
                # 1. Assegnazione Vincoli/Eccezioni
                if (nome, g) in st.session_state.eccezioni:
                    matrice_turni[nome][str_g] = st.session_state.eccezioni[(nome, g)]
                elif g == 1:
                    t_in = turni_inizio_dict.get(nome, "M")
                    if esoneri_dict.get(nome, False) and t_in == "N":
                        t_in = "R"
                    matrice_turni[nome][str_g] = t_in
                else:
                    prev = matrice_turni[nome].get(str(g - 1), "")
                    if prev in ["N", "N_DES"]:
                        matrice_turni[nome][str_g] = "S"  # N -> S tassativo
                    elif prev == "S":
                        matrice_turni[nome][str_g] = "R"  # S -> R tassativo
                    elif "Mattina" in tf:
                        giorni_lav = sum(1 for d in range(max(1, g-6), g) if matrice_turni[nome].get(str(d)) in ["M", "PC"])
                        matrice_turni[nome][str_g] = "R" if giorni_lav >= 5 else "M"
                    elif "Pomeriggio" in tf:
                        giorni_lav = sum(1 for d in range(max(1, g-6), g) if matrice_turni[nome].get(str(d)) in ["P", "CFD"])
                        matrice_turni[nome][str_g] = "R" if giorni_lav >= 5 else "P"
                    else:
                        # Segue la rotazione ciclica sfalsata M-P-N-S-R
                        t_teorico = rotazione[(g - 1 + offset) % len(rotazione)]
                        if t_teorico == "N" and esoneri_dict.get(nome, False):
                            t_teorico = "R"
                        if t_teorico == "M" and prev in ["P", "P_DES"]:  # Blocco 11h
                            t_teorico = "P"
                        matrice_turni[nome][str_g] = t_teorico

        # Step 2: Bilanciamento Intelligente per Raggiungere le 156h Senza Intaccare il Fabbisogno
        for _ in range(25):
            ore_per_inf = {
                n: sum(pesi_ore.get(matrice_turni[n].get(str(g), "R"), 0) for g in range(1, giorni_nel_mese + 1))
                for n in nomi_inf_list
            }
            sotto = [n for n, h in ore_per_inf.items() if h < 152]
            sopra = [n for n, h in ore_per_inf.items() if h > 162]
            
            if not sotto or not sopra:
                break
                
            scambio_fatto = False
            for n_under in sotto:
                for n_over in sopra:
                    for g in range(1, giorni_nel_mese + 1):
                        str_g = str(g)
                        t_u = matrice_turni[n_under].get(str_g)
                        t_o = matrice_turni[n_over].get(str_g)
                        
                        if t_u == 'R' and t_o in ['M', 'P', 'CFD', 'PC']:
                            ok = True
                            if t_o in ['M', 'PC'] and g > 1:
                                prev_u = matrice_turni[n_under].get(str(g-1), '')
                                if prev_u in ['P', 'N', 'S']:
                                    ok = False
                            tf_u = turni_fissi_dict.get(n_under, 'Nessuno')
                            if 'Mattina' in tf_u and t_o not in ['M', 'PC']:
                                ok = False
                            if 'Pomeriggio' in tf_u and t_o not in ['P', 'CFD']:
                                ok = False
                            if (n_under, g) in st.session_state.eccezioni or (n_over, g) in st.session_state.eccezioni:
                                ok = False
                                
                            if ok:
                                matrice_turni[n_under][str_g] = t_o
                                matrice_turni[n_over][str_g] = 'R'
                                scambio_fatto = True
                                break
                    if scambio_fatto:
                        break
                if scambio_fatto:
                    break
            if not scambio_fatto:
                break

        # Costruzione DataFrame finale
        dati_matrice = []
        for nome in nomi_inf_list:
            riga = {"INFERMIERE": nome}
            ore_totali = 0
            for g in range(1, giorni_nel_mese + 1):
                t_val = matrice_turni[nome].get(str(g), "R")
                t_disp = t_val.replace("_DES", "")
                riga[str(g)] = t_disp
                ore_totali += pesi_ore.get(t_val, 0)
                
            riga["TOT ore"] = ore_totali
            dati_matrice.append(riga)

        df_matrice = pd.DataFrame(dati_matrice).set_index("INFERMIERE")

        def applica_stile(val):
            colori = pd.DataFrame('', index=df_matrice.index, columns=df_matrice.columns)
            for col_dom in colonne_domeniche:
                if col_dom in colori.columns:
                    colori[col_dom] = 'background-color: #FFFF99; font-weight: bold;'
            return colori

        st.dataframe(df_matrice.style.apply(applica_stile, axis=None), use_container_width=True, height=480)

        # --- EXPORT CSV ---
        csv_data = df_matrice.to_csv().encode('utf-8')
        st.download_button(
            label="📥 Scarica Tabella Turni (.csv per Excel)",
            data=csv_data,
            file_name=f"Turnistica_Infermieri_{MESI_IT[mese]}_{anno}.csv",
            mime="text/csv",
            type="primary"
        )

        # --- CONTEGGIO CON RIGHE CONTA M, CONTA P, CONTA N (STILE MATRICE UFFICIALE) ---
        st.markdown("---")
        st.markdown("##### 📈 Riepilogo Giornaliero Presenze (Righe CONTA M / CONTA P / CONTA N)")
        
        tot_m_row = {"Turno / Controllo": f"CONTA M (Fabbisogno: {req_m})"}
        tot_p_row = {"Turno / Controllo": f"CONTA P (Fabbisogno: {req_p})"}
        tot_n_row = {"Turno / Controllo": f"CONTA N (Fabbisogno: {req_n})"}
        
        for g in range(1, giorni_nel_mese + 1):
            col_g = [riga[str(g)] for riga in dati_matrice]
            tot_m_row[str(g)] = col_g.count("M") + col_g.count("PC")
            tot_p_row[str(g)] = col_g.count("P") + col_g.count("CFD")
            tot_n_row[str(g)] = col_g.count("N")
            
        tot_m_row["TOT ore"] = ""
        tot_p_row["TOT ore"] = ""
        tot_n_row["TOT ore"] = ""

        df_riepilogo = pd.DataFrame([tot_m_row, tot_p_row, tot_n_row]).set_index("Turno / Controllo")
        st.dataframe(df_riepilogo, use_container_width=True)