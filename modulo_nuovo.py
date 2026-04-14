import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from openpyxl import load_workbook

# --- FUNZIONE DATA FORMATO ESTESO (ES: 12 APRILE 2026) ---
def formatta_data_it(data_obj):
    mesi = {
        1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE",
        5: "MAGGIO", 6: "GIUGNO", 7: "LUGLIO", 8: "AGOSTO",
        9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
    }
    return f"{data_obj.day} {mesi[data_obj.month]} {data_obj.year}"

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Generatore Modulo", layout="wide")

# TITOLO MODIFICATO: Modulo Ordine xlsx in grassetto e "BY_EmmeKappa" piccolo e corsivo
st.markdown(f"### Modulo Ordine xlsx <span style='font-size: 0.6em; font-style: italic;'>by_EmmeKappa</span>", unsafe_allow_html=True)
st.header("Generatore Modulo")

# --- CSS ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    .stDownloadButton>button { width: 100%; border-radius: 8px; background-color: #28a745 !important; color: white !important; }
    input:placeholder-shown, textarea:placeholder-shown { border: 1.5px solid #ff4b4b !important; }
    input:not(:placeholder-shown), textarea:not(:placeholder-shown) { border: 1.5px solid #28a745 !important; background-color: #f0fff4 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- CONFIGURAZIONE FILE ---
MODELLO_EXCEL = "modello.xlsx" 
OUTPUT_FOLDER = "Moduli_Finali"
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

if 'lista_articoli' not in st.session_state:
    st.session_state.lista_articoli = []
if 'form_reset' not in st.session_state:
    st.session_state.form_reset = 0

PRODOTTI = [
    "Select...", "Fanali C L 201", "Fanali C L 301", "Fanali C L 401", "Fanali CL 351",
    "INTERNO COMPLETO CLR 70", "INTERNO COMPLETO CLR 70H", 
    "Fanali CS 155", "Fanali CS 250", "Fanali CL 299", 
    "Fanali CL 299 H", "Fanali CL 301 H", "Fanali CL 401 H", 
    "Fotocellula", "Regolatore_di_Carica", 
    "Alimentatore 220VAC/ 12VCC 1A", "_Batteria", "Pannello Solare", "Lampade_12V."
]
# COLORE LUCE MODIFICATO: Parte con Select...
COLORI_LUCE = ["Select...", "Bianca", "Rossa", "Verde", "Gialla", "Settorizzata", "Oscurato"]
GPS_OPZIONI = ["NO", "SI"]

st.divider()

# --- 1. DATI ORDINE ---
st.subheader("Dati Ordine")
col_test1, col_test2 = st.columns(2)
with col_test1:
    cliente = st.text_input("Cliente", key="fisso_cliente").upper()
with col_test2:
    commessa = st.text_input("COMMESSA N°", key="fisso_commessa").upper()

col_test3, col_test4 = st.columns(2)
with col_test3:
    data_consegna = st.date_input("SPEDIRE ENTRO IL", value=date.today(), format="DD/MM/YYYY", key="fisso_consegna")
    consegna_str = formatta_data_it(data_consegna)
with col_test4:
    dest_scelto = st.radio("DESTINAZIONE MERCE", ["CLIENTE", "RESINEX"], horizontal=True, key="fisso_dest")

st.divider()

# --- 2. AGGIUNGI ARTICOLO ---
st.subheader("Seleziona_articolo")
suffix = st.session_state.form_reset

tipo_int = st.radio("Lavorazione per", ["VENDITA", "RIPARAZIONE"], horizontal=True, key=f"tipo_{suffix}")
prod_scelto = st.selectbox("Tipologia Articolo", PRODOTTI, key=f"prod_{suffix}")

d_art = {"lamp": "", "port": "", "colore": "", "gps": "", "note": ""}
qta = 1

if prod_scelto != "Select...":
    is_fanaleria = any(x in prod_scelto for x in ["Fanali", "INTERNO"])
    if is_fanaleria:
        ca, cb, cc, cd = st.columns(4)
        with ca: d_art["colore"] = st.selectbox("Colore Luce", COLORI_LUCE, index=0, key=f"col_{suffix}")
        with cb: qta = st.number_input("Q.TA'", min_value=1, step=1, key=f"qta_{suffix}")
        with cc: d_art["port"] = st.text_input("Portata M_N_", key=f"port_{suffix}")
        with cd: d_art["gps"] = st.radio("Sincro GPS", GPS_OPZIONI, horizontal=True, key=f"gps_{suffix}")
        d_art["lamp"] = st.text_area("Caratteristica di lampeggio", key=f"lamp_{suffix}")
        d_art["note"] = st.text_area("NOTE / SPECIFICHE TECNICHE  ", key=f"note_{suffix}")
    else:
        qta = st.number_input("Q.TA'", min_value=1, step=1, key=f"qta_gen_{suffix}")
        d_art["note"] = st.text_area("NOTE / SPECIFICHE TECNICHE  ", key=f"note_gen_{suffix}")
        d_art["gps"] = ""

    # CONTROLLO CAMPI VUOTI (Incluso Colore)
    campi_vuoti = is_fanaleria and (not d_art["port"] or not d_art["lamp"] or d_art["colore"] == "Selezionare...")
    
    if campi_vuoti:
        st.warning("⚠️ Attenzione: Portata, Lampeggio o Colore non compilati.")
        if st.checkbox("Confermo di voler inserire l'articolo incompleto", key=f"chk_{suffix}"):
            if st.button("➕ AGGIUNGI ARTICOLO (INCOMPLETO)"):
                if cliente and commessa:
                    st.session_state.lista_articoli.append({
                        "tipo": tipo_int, "prod": prod_scelto, "qta": qta,
                        "lamp": d_art["lamp"].upper(), "port": d_art["port"].upper(),
                        "colore": d_art["colore"].upper(), "gps": d_art["gps"],
                        "note": d_art["note"].upper()
                    })
                    st.session_state.form_reset += 1 
                    st.rerun()
                else:
                    st.error("Mancano Cliente o Commessa!")
    else:
        if st.button("➕ AGGIUNGI ARTICOLO"):
            if cliente and commessa:
                st.session_state.lista_articoli.append({
                    "tipo": tipo_int, "prod": prod_scelto, "qta": qta,
                    "lamp": d_art["lamp"].upper(), "port": d_art["port"].upper(),
                    "colore": d_art["colore"].upper(), "gps": d_art["gps"],
                    "note": d_art["note"].upper()
                })
                st.session_state.form_reset += 1 
                st.rerun()
            else:
                st.error("Mancano Cliente o Commessa!")

# --- 3. RIEPILOGO E SALVATAGGIO ---
if st.session_state.lista_articoli:
    st.divider()
    st.subheader(f"📋 Articoli inseriti ({len(st.session_state.lista_articoli)})")
    st.dataframe(pd.DataFrame(st.session_state.lista_articoli), use_container_width=True, hide_index=True)

    if st.button("🚀 GENERA FILE EXCEL"):
        try:
            wb = load_workbook(MODELLO_EXCEL)
            ws = wb.active 
            ws["D4"], ws["D5"], ws["D42"], ws["D44"] = cliente, commessa, consegna_str, dest_scelto
             # ws["E46"] = f"DATA EMISSIONE {formatta_data_it(datetime.now())}"
            
            for i, art in enumerate(st.session_state.lista_articoli[:6]):
                r_label = 7 if i == 0 else (20 if i == 1 else (31 if i == 2 else (52 if i == 3 else (65 if i == 4 else 76))))
                r_base = r_label + 3
                ws[f"D{r_label}"], ws[f"D{r_base}"], ws[f"D{r_base+1}"] = art["tipo"], art["prod"], art["qta"]
                ws[f"D{r_base+2}"], ws[f"D{r_base+3}"], ws[f"D{r_base+4}"] = art["lamp"], art["port"], art["colore"]
                ws[f"D{r_base+5}"], ws[f"D{r_base+6}"] = art["gps"], art["note"] 
            
            # NOME FILE DINAMICO: Commessa_Numero_NomeCliente
            nome_pulito_cliente = cliente.replace(" ", "_")
            nome_file_dinamico = f"Commessa_{commessa}_{nome_pulito_cliente}.xlsx"
            
            path = f"{OUTPUT_FOLDER}/{nome_file_dinamico}"
            wb.save(path)
            st.success(f"✅ Excel generato: {nome_file_dinamico}")
            with open(path, "rb") as f:
                st.download_button("📥 SCARICA ORA IL FILE EXCEL", f, file_name=nome_file_dinamico)
        except Exception as e: st.error(f"Errore: {e}")

    if st.button("🗑️ CANCELLA TUTTA LA LISTA"):
        st.session_state.lista_articoli = []
        st.session_state.form_reset = 0
        st.rerun()
