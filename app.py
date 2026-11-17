import streamlit as st
import pandas as pd
import calendar
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Turnistica Mensile Infermieri", layout="wide")

MESI_IT = {
    1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE",
    5: "MAGGIO", 6: "GIUGNO", 7: "LUGLIO", 8: "AGOSTO",
    9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
}

st.title("Pianificatore Turni Infermieristici")

# --- SIDEBAR: CONFIGURAZIONE MESE E REPARTO ---
st.sidebar.header("1. Impostazioni Mese e Reparto")
intestazione_dip = st.sidebar.text_input("Intestazione Dipartimento:", value="")
reparto = st.sidebar.text_input("Nome Reparto:", value="")
anno = st.sidebar.number_input("Anno", min_value=2024, max_value=2030, value=2026)
mese = st.sidebar.selectbox("Mese", list(range(1, 13)), index=10, format_func=lambda x: MESI_IT[x])
data_emissione = st.sidebar.text_input("Data Emissione:", value="")

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

# Inizializzazione Session State per il Personale
if 'df_pers' not in st.session_state or len(st.session_state.df_pers) != num_infermieri:
    dati_iniziali = []
    for idx in range(num_infermieri):
        dati_iniziali.append({
            "Nome": f"Infermiere {idx+1}",
            "Qualifica": "Infermiere",
            "Esonero_Notte": "No",
            "Turno_Fisso": "Nessuno",
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
    st.info("💡 Imposta liberamente esoneri notturni, turni fissi e il turno iniziale del mese. Clicca su **Salva Modifiche**.")
    
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
                    options=["M", "P", "N", "S", "R", "FE", "104", "ASP", "MAL", "RO", "CFD", "PC", "PR"],
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
                "M (Mattina)", "P (Pomeriggio)", "N (Notte)", "S (Smonto)", "R (Riposo)",
                "FE (Ferie)", "PR (Permesso)", "104 (Legge 104)", "ASP (Aspettativa)", "MAL (Malattia)",
                "RO (Recupero Ore)", "CFD (Centro Ferite Difficili - P)", "PC (Prericovero Centralizzato - M)",
                "M_DES (Desiderata Mattina)", "P_DES (Desiderata Pomeriggio)", "N_DES (Desiderata Notte)"
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

# --- FUNZIONE GENERAZIONE FILE EXCEL COLORATO ---
def genera_excel_colorato(dati_matrice, matrice_turni_raw, giorni_settimana, colonne_domeniche, tot_m_row, tot_p_row, tot_n_row):
    wb = Workbook()
    ws = wb.active
    ws.title = f"Turni_{MESI_IT[mese]}"

    thin_side = Side(border_style="thin", color="000000")
    med_side = Side(border_style="medium", color="000000")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_box = Border(left=med_side, right=med_side, top=med_side, bottom=med_side)

    font_header = Font(bold=True, size=11)
    font_bold = Font(bold=True, size=10)
    font_normal = Font(size=10)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    fill_header_grey = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
    fill_domenica = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    
    fill_fe = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    font_fe = Font(bold=True, color="155724", size=10)

    fill_asp_pr = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    font_asp_pr = Font(bold=True, color="004085", size=10)

    fill_mal_ro = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    font_mal_ro = Font(bold=True, color="856404", size=10)

    fill_des = PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid")
    font_des = Font(bold=True, color="3B1B5B", size=10)

    tot_colonne = giorni_nel_mese + 2

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
    ws.cell(1, 1, intestazione_dip.upper()).font = font_header

    ws.merge_cells(start_row=1, start_column=15, end_row=1, end_column=24)
    ws.cell(1, 15, f"DATA EMISSIONE: {data_emissione}").font = font_header

    ws.merge_cells(start_row=1, start_column=25, end_row=1, end_column=tot_colonne)
    ws.cell(1, 25, "ORE DOVUTE: 156").font = font_header

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=14)
    ws.cell(2, 1, reparto.upper()).font = font_header

    ws.merge_cells(start_row=2, start_column=15, end_row=2, end_column=tot_colonne)
    ws.cell(2, 15, f"MESE DI {MESI_IT[mese]} {anno}").font = font_header

    for r in [1, 2]:
        for c in range(1, tot_colonne + 1):
            cell = ws.cell(r, c)
            cell.fill = fill_header_grey
            cell.border = border_box
            cell.alignment = align_center

    ws.cell(4, 1, "GIORNO").font = font_bold
    ws.cell(4, 1).alignment = align_center
    ws.cell(4, 1).border = border_cell

    ws.cell(5, 1, "SETTIMANA").font = font_bold
    ws.cell(5, 1).alignment = align_center
    ws.cell(5, 1).border = border_cell

    for g in range(1, giorni_nel_mese + 1):
        col_idx = g + 1
        c_num = ws.cell(4, col_idx, g)
        c_let = ws.cell(5, col_idx, giorni_settimana[g - 1])
        
        c_num.font = font_bold
        c_let.font = font_bold
        c_num.alignment = align_center
        c_let.alignment = align_center
        c_num.border = border_cell
        c_let.border = border_cell

        if str(g) in colonne_domeniche:
            c_num.fill = fill_domenica
            c_let.fill = fill_domenica

    ws.cell(4, tot_colonne, "").border = border_cell
    c_tot_h = ws.cell(5, tot_colonne, "TOT ore")
    c_tot_h.font = font_bold
    c_tot_h.alignment = align_center
    c_tot_h.border = border_cell

    curr_row = 6
    for riga_dati in dati_matrice:
        nome = riga_dati["INFERMIERE"]
        c_nome = ws.cell(curr_row, 1, nome)
        c_nome.font = font_bold
        c_nome.alignment = align_left
        c_nome.border = border_cell

        for g in range(1, giorni_nel_mese + 1):
            col_idx = g + 1
            val_disp = riga_dati[str(g)]
            val_raw = matrice_turni_raw[nome].get(str(g), "R")

            cell = ws.cell(curr_row, col_idx, val_disp)
            cell.alignment = align_center
            cell.border = border_cell
            cell.font = font_normal

            if str(g) in colonne_domeniche:
                cell.fill = fill_domenica

            if val_disp == "FE":
                cell.fill = fill_fe
                cell.font = font_fe
            elif val_disp in ["ASP", "PR"]:
                cell.fill = fill_asp_pr
                cell.font = font_asp_pr
            elif val_disp in ["MAL", "RO"]:
                cell.fill = fill_mal_ro
                cell.font = font_mal_ro
            elif "DES" in val_raw:
                cell.fill = fill_des
                cell.font = font_des

        c_tot = ws.cell(curr_row, tot_colonne, riga_dati["TOT ore"])
        c_tot.font = font_bold
        c_tot.alignment = align_center
        c_tot.border = border_cell

        curr_row += 1

    curr_row += 1
    for r_conta in [tot_m_row, tot_p_row, tot_n_row]:
        c_label = ws.cell(curr_row, 1, r_conta["INFERMIERE"])
        c_label.font = font_bold
        c_label.alignment = align_left
        c_label.fill = fill_header_grey
        c_label.border = border_cell

        for g in range(1, giorni_nel_mese + 1):
            col_idx = g + 1
            c_val = ws.cell(curr_row, col_idx, r_conta[str(g)])
            c_val.font = font_bold
            c_val.alignment = align_center
            c_val.border = border_cell
            if str(g) in colonne_domeniche:
                c_val.fill = fill_domenica

        ws.cell(curr_row, tot_colonne, "").border = border_cell
        curr_row += 1

    ws.column_dimensions["A"].width = 28
    for g in range(1, giorni_nel_mese + 1):
        ws.column_dimensions[get_column_letter(g + 1)].width = 5.5
    ws.column_dimensions[get_column_letter(tot_colonne)].width = 10

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# --- FUNZIONI AUSILIARIE PER VINCOLI CCNL SANITÀ ---
def conta_giorni_consecutivi_senza_riposo(matrice_inf, giorno_attuale):
    """Conta quanti giorni consecutivi l'operatore ha lavorato (incluso lo smonto) prima del giorno attuale."""
    consecutivi = 0
    for d in range(giorno_attuale - 1, 0, -1):
        t = matrice_inf.get(str(d), "R")
        # R, FE, PR, 104, ASP, MAL, RO interrompono la sequenza di servizio attivo continuo
        if t in ["R", "FE", "PR", "104", "ASP", "MAL", "RO"]:
            break
        consecutivi += 1
    return consecutivi

def passaggio_valido_ccnl(prev_turno, nuovo_turno, giorni_consecutivi, ha_esonero_notte, turno_fisso):
    """Verifica il rispetto delle 11h di riposo, smonto post-notte e tetto max 6 giorni consecutivi."""
    # 1. Tetto massimo 6 giorni consecutivi di servizio (al 7° giorno scatta il riposo obbligatorio CCNL)
    if nuovo_turno not in ["R", "FE", "PR", "104", "ASP", "MAL", "RO"] and giorni_consecutivi >= 6:
        return False

    # 2. Esonero Notte
    if ha_esonero_notte and nuovo_turno in ["N", "N_DES"]:
        return False

    # 3. Turno Fisso
    if "Mattina" in turno_fisso and nuovo_turno not in ["M", "PC", "R", "FE", "PR", "104", "ASP", "MAL", "RO"]:
        return False
    if "Pomeriggio" in turno_fisso and nuovo_turno not in ["P", "CFD", "R", "FE", "PR", "104", "ASP", "MAL", "RO"]:
        return False

    # 4. Regola Smonto: 'S' esiste SOLO se ieri era Notte ('N')
    if prev_turno in ["N", "N_DES"] and nuovo_turno != "S":
        return False
    if nuovo_turno == "S" and prev_turno not in ["N", "N_DES"]:
        return False

    # 5. Riposo minimo giornaliero di 11 ore (CCNL Sanità):
    # Vietato fare Mattina (07:00) dopo Pomeriggio (finisce alle 21:00 -> solo 10h di stacco) o dopo Notte
    if prev_turno in ["P", "CFD", "P_DES"] and nuovo_turno in ["M", "PC", "M_DES"]:
        return False
    # Dopo lo Smonto ('S', che termina alle 07:00 del mattino), è ammesso riprendere con 'M' o 'P' (24h/31h di stacco) oppure 'R'
    return True


# --- TAB 3: GENERATORE MATRICE AZIENDALE VALIDATA ---
with tab_matrice:
    st.markdown("""
        <style>
        .header-box { border: 2px solid #000; padding: 6px; background-color: #f9f9f9; text-align: center; font-weight: bold; }
        .header-title { font-size: 15px; text-transform: uppercase; }
        .sub-header { border-left: 2px solid #000; border-right: 2px solid #000; border-bottom: 2px solid #000; padding: 4px; font-weight: bold; background-color: #f0f0f0; }
        </style>
    """, unsafe_allow_html=True)

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
            "FE": 6, "PR": 6, "104": 6, "ASP": 6, "MAL": 6, "RO": 6,
            "R": 0
        }

        # --- PASSAGGIO 1: ROTAZIONE CCNL CON RIENTRO POST-SMONTO E RIPOSO OGNI MAX 6 GG ---
        # Ciclo base ampliato che permette la sequenza M -> P -> N -> S -> M -> P -> R (6 gg lavorati + 1 riposo)
        ciclo_ccnl = ["M", "P", "N", "S", "M", "P", "R"]
        
        for idx, nome in enumerate(nomi_inf_list):
            offset = (idx * 2) % len(ciclo_ccnl)
            es_notte = esoneri_dict.get(nome, False)
            tf = turni_fissi_dict.get(nome, "Nessuno")
            
            for g in range(1, giorni_nel_mese + 1):
                str_g = str(g)
                consecutivi = conta_giorni_consecutivi_senza_riposo(matrice_turni[nome], g)
                
                if (nome, g) in st.session_state.eccezioni:
                    matrice_turni[nome][str_g] = st.session_state.eccezioni[(nome, g)]
                elif g == 1:
                    t_in = turni_inizio_dict.get(nome, "M")
                    if es_notte and t_in in ["N", "N_DES"]:
                        t_in = "R"
                    matrice_turni[nome][str_g] = t_in
                else:
                    prev = matrice_turni[nome].get(str(g - 1), "")
                    
                    if prev in ["N", "N_DES"]:
                        matrice_turni[nome][str_g] = "S"  # Dopo la notte c'è SEMPRE lo smonto
                    elif consecutivi >= 6:
                        matrice_turni[nome][str_g] = "R"  # Al 7° giorno scatta il riposo obbligatorio CCNL
                    elif "Mattina" in tf:
                        matrice_turni[nome][str_g] = "R" if consecutivi >= 5 else "M"
                    elif "Pomeriggio" in tf:
                        matrice_turni[nome][str_g] = "R" if consecutivi >= 5 else "P"
                    else:
                        t_teorico = ciclo_ccnl[(g - 1 + offset) % len(ciclo_ccnl)]
                        # Se il ciclo propone 'S' ma ieri non era notte, diventa 'R'
                        if t_teorico == "S":
                            t_teorico = "R"
                        if not passaggio_valido_ccnl(prev, t_teorico, consecutivi, es_notte, tf):
                            # Ripiego sicuro conforme al CCNL
                            if passaggio_valido_ccnl(prev, "P", consecutivi, es_notte, tf):
                                t_teorico = "P"
                            elif passaggio_valido_ccnl(prev, "M", consecutivi, es_notte, tf):
                                t_teorico = "M"
                            else:
                                t_teorico = "R"
                        matrice_turni[nome][str_g] = t_teorico

        # --- PASSAGGIO 2: COPERTURA RIGOROSA DEL FABBISOGNO GIORNALIERO (SALTANDO RIPOSI SE NECESSARIO) ---
        for g in range(1, giorni_nel_mese + 1):
            str_g = str(g)
            
            # 1. Bilanciamento Notti (N)
            attuali_n = sum(1 for n in nomi_inf_list if matrice_turni[n].get(str_g) in ["N", "N_DES"])
            if attuali_n < req_n:
                # Ordiniamo i candidati dando priorità a chi ha meno ore accumulate finora
                candidati = [
                    n for n in nomi_inf_list
                    if (n, g) not in st.session_state.eccezioni
                    and (g == giorni_nel_mese or (n, g + 1) not in st.session_state.eccezioni)
                    and matrice_turni[n].get(str_g) in ["R", "M", "P"]
                    and passaggio_valido_ccnl(
                        matrice_turni[n].get(str(g-1), "") if g > 1 else "",
                        "N",
                        conta_giorni_consecutivi_senza_riposo(matrice_turni[n], g),
                        esoneri_dict.get(n, False),
                        turni_fissi_dict.get(n, "Nessuno")
                    )
                ]
                candidati.sort(key=lambda x: sum(pesi_ore.get(matrice_turni[x].get(str(d), "R"), 0) for d in range(1, g)))
                for n in candidati[:(req_n - attuali_n)]:
                    matrice_turni[n][str_g] = "N"
                    if g < giorni_nel_mese:
                        matrice_turni[n][str(g + 1)] = "S"
            elif attuali_n > req_n:
                candidati = [n for n in nomi_inf_list if matrice_turni[n].get(str_g) == "N" and (n, g) not in st.session_state.eccezioni]
                for n in candidati[:(attuali_n - req_n)]:
                    matrice_turni[n][str_g] = "R"
                    if g < giorni_nel_mese and matrice_turni[n].get(str(g + 1)) == "S":
                        matrice_turni[n][str(g + 1)] = "R"

            # 2. Bilanciamento Mattine (M) - Può convertire 'R' in 'M' anche il giorno dopo uno smonto ('S')
            attuali_m = sum(1 for n in nomi_inf_list if matrice_turni[n].get(str_g) in ["M", "PC", "M_DES"])
            if attuali_m < req_m:
                candidati = [
                    n for n in nomi_inf_list
                    if (n, g) not in st.session_state.eccezioni
                    and matrice_turni[n].get(str_g) in ["R", "P"]
                    and (g == giorni_nel_mese or matrice_turni[n].get(str(g + 1)) != "S")
                    and passaggio_valido_ccnl(
                        matrice_turni[n].get(str(g-1), "") if g > 1 else "",
                        "M",
                        conta_giorni_consecutivi_senza_riposo(matrice_turni[n], g),
                        esoneri_dict.get(n, False),
                        turni_fissi_dict.get(n, "Nessuno")
                    )
                ]
                candidati.sort(key=lambda x: sum(pesi_ore.get(matrice_turni[x].get(str(d), "R"), 0) for d in range(1, g)))
                for n in candidati[:(req_m - attuali_m)]:
                    matrice_turni[n][str_g] = "M"
            elif attuali_m > req_m:
                candidati = [n for n in nomi_inf_list if matrice_turni[n].get(str_g) == "M" and (n, g) not in st.session_state.eccezioni]
                candidati.sort(key=lambda x: sum(pesi_ore.get(matrice_turni[x].get(str(d), "R"), 0) for d in range(1, g)), reverse=True)
                for n in candidati[:(attuali_m - req_m)]:
                    matrice_turni[n][str_g] = "R"

            # 3. Bilanciamento Pomeriggi (P) - Può convertire 'R' in 'P' anche il giorno dopo uno smonto ('S')
            attuali_p = sum(1 for n in nomi_inf_list if matrice_turni[n].get(str_g) in ["P", "CFD", "P_DES"])
            if attuali_p < req_p:
                candidati = [
                    n for n in nomi_inf_list
                    if (n, g) not in st.session_state.eccezioni
                    and matrice_turni[n].get(str_g) == "R"
                    and (g == giorni_nel_mese or matrice_turni[n].get(str(g + 1)) != "S")
                    and passaggio_valido_ccnl(
                        matrice_turni[n].get(str(g-1), "") if g > 1 else "",
                        "P",
                        conta_giorni_consecutivi_senza_riposo(matrice_turni[n], g),
                        esoneri_dict.get(n, False),
                        turni_fissi_dict.get(n, "Nessuno")
                    )
                ]
                candidati.sort(key=lambda x: sum(pesi_ore.get(matrice_turni[x].get(str(d), "R"), 0) for d in range(1, g)))
                for n in candidati[:(req_p - attuali_p)]:
                    matrice_turni[n][str_g] = "P"
            elif attuali_p > req_p:
                candidati = [n for n in nomi_inf_list if matrice_turni[n].get(str_g) == "P" and (n, g) not in st.session_state.eccezioni]
                candidati.sort(key=lambda x: sum(pesi_ore.get(matrice_turni[x].get(str(d), "R"), 0) for d in range(1, g)), reverse=True)
                for n in candidati[:(attuali_p - req_p)]:
                    matrice_turni[n][str_g] = "R"

        # --- PASSAGGIO 3: EQUILIBRATURA ORE VERSO LE 156H SENZA ALTERARE IL FABBISOGNO ---
        for _ in range(30):
            ore_per_inf = {
                n: sum(pesi_ore.get(matrice_turni[n].get(str(g), "R"), 0) for g in range(1, giorni_nel_mese + 1))
                for n in nomi_inf_list
            }
            sotto = sorted([n for n, h in ore_per_inf.items() if h < 154], key=lambda x: ore_per_inf[x])
            sopra = sorted([n for n, h in ore_per_inf.items() if h > 160], key=lambda x: ore_per_inf[x], reverse=True)
            
            if not sotto or not sopra:
                break
                
            scambio_fatto = False
            for n_under in sotto:
                for n_over in sopra:
                    for g in range(1, giorni_nel_mese + 1):
                        str_g = str(g)
                        t_u = matrice_turni[n_under].get(str_g)
                        t_o = matrice_turni[n_over].get(str_g)
                        
                        # Scambio 1-a-1 tra chi è in Riposo ('R') e chi ha un turno diurno ('M' o 'P'):
                        # Così il numero totale di infermieri in servizio nel giorno resta IDENTICO!
                        if t_u == 'R' and t_o in ['M', 'P']:
                            if (n_under, g) in st.session_state.eccezioni or (n_over, g) in st.session_state.eccezioni:
                                continue
                            
                            prev_u = matrice_turni[n_under].get(str(g - 1), "") if g > 1 else ""
                            next_u = matrice_turni[n_under].get(str(g + 1), "") if g < giorni_nel_mese else ""
                            consec_u = conta_giorni_consecutivi_senza_riposo(matrice_turni[n_under], g)
                            
                            # Verifica vincoli CCNL su n_under sia verso ieri che verso domani
                            valido_ieri = passaggio_valido_ccnl(
                                prev_u, t_o, consec_u, esoneri_dict.get(n_under, False), turni_fissi_dict.get(n_under, "Nessuno")
                            )
                            valido_domani = True
                            if next_u != "":
                                valido_domani = passaggio_valido_ccnl(
                                    t_o, next_u, consec_u + 1, esoneri_dict.get(n_under, False), turni_fissi_dict.get(n_under, "Nessuno")
                                )
                                
                            if valido_ieri and valido_domani:
                                matrice_turni[n_under][str_g] = t_o
                                matrice_turni[n_over][str_g] = 'R'
                                scambio_fatto = True
                                break
                    if scambio_fatto: break
                if scambio_fatto: break
            if not scambio_fatto: break

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

        def colora_griglia(df):
            stili = pd.DataFrame('', index=df.index, columns=df.columns)
            
            for col_dom in colonne_domeniche:
                if col_dom in stili.columns:
                    stili[col_dom] = 'background-color: #FFFF99;'
                    
            for r in df.index:
                for c in df.columns:
                    val_disp = str(df.loc[r, c])
                    val_raw = matrice_turni[r].get(c, "") if c != "TOT ore" else ""
                    if val_disp == 'FE':
                        stili.loc[r, c] = 'background-color: #D4EDDA; color: #155724; font-weight: bold;'
                    elif val_disp in ['ASP', 'PR']:
                        stili.loc[r, c] = 'background-color: #CCE5FF; color: #004085; font-weight: bold;'
                    elif val_disp in ['MAL', 'RO']:
                        stili.loc[r, c] = 'background-color: #FFF3CD; color: #856404; font-weight: bold;'
                    elif 'DES' in val_raw:
                        stili.loc[r, c] = 'background-color: #E2D9F3; color: #3B1B5B; font-weight: bold;'
            return stili

        st.dataframe(df_matrice.style.apply(colora_griglia, axis=None), use_container_width=True, height=480)

        # --- CALCOLO RIGHE CONTA M, P, N (ALLINEATE AL FABBISOGNO) ---
        tot_m_row = {"INFERMIERE": f"CONTA M (Fabbisogno: {req_m})"}
        tot_p_row = {"INFERMIERE": f"CONTA P (Fabbisogno: {req_p})"}
        tot_n_row = {"INFERMIERE": f"CONTA N (Fabbisogno: {req_n})"}
        
        for g in range(1, giorni_nel_mese + 1):
            col_g = [riga[str(g)] for riga in dati_matrice]
            tot_m_row[str(g)] = col_g.count("M") + col_g.count("PC")
            tot_p_row[str(g)] = col_g.count("P") + col_g.count("CFD")
            tot_n_row[str(g)] = col_g.count("N")
            
        tot_m_row["TOT ore"] = ""
        tot_p_row["TOT ore"] = ""
        tot_n_row["TOT ore"] = ""

        # --- EXPORT EXCEL (.XLSX) ---
        excel_bytes = genera_excel_colorato(
            dati_matrice, matrice_turni, giorni_settimana, colonne_domeniche, tot_m_row, tot_p_row, tot_n_row
        )
        nome_rep_file = reparto.replace(' ', '_') if reparto.strip() else "Reparto"

        st.download_button(
            label="📥 Scarica Report Turni a Colori (.xlsx per Excel)",
            data=excel_bytes,
            file_name=f"Report_Turni_{nome_rep_file}_{MESI_IT[mese]}_{anno}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        # --- CONTEGGIO A SCHERMO ---
        st.markdown("---")
        st.markdown("##### 📈 Riepilogo Giornaliero Presenze (Righe CONTA M / CONTA P / CONTA N)")
        
        df_riepilogo = pd.DataFrame([tot_m_row, tot_p_row, tot_n_row]).set_index("INFERMIERE")
        st.dataframe(df_riepilogo, use_container_width=True)