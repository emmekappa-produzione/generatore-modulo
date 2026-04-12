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

# --- CSS PER CORNICI ROSSE (VUOTE) E VERDI (PIENE) ---
st.markdown("""
    <style>
    /* Campi Vuoti (mostrano il segnaposto) */
    input:placeholder-shown, textarea:placeholder-shown {
        border: 2px solid #FF4B4B !important;
    }
    /* Campi Pieni (non mostrano più il segnaposto) */
    input:not(:placeholder-shown), textarea:not(:placeholder-shown) {
        border: 2px solid #28A745 !important;
        background-color: #F0FFF4 !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Generatore Modulo")
st.header("Generatore Modulo Produzione")

# --- CONFIGURAZIONE FILE ---
MODELLO_EXCEL = "modello.xlsx" 
OUTPUT_FOLDER = "Moduli_Finali"
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# Inizializzazione stati
if 'lista_articoli' not in st.session_state:
    st.session_state.lista_articoli = []
if 'form_reset' not in st.session_state:
    st.session_state.form_reset = 0

# --- LISTA PRODOTTI ORIGINALE ---
PRODOTTI = [
    "Selezionare...", "Fanali C L 201", "Fanali C L 301", "Fanali C L 401", 
    "INTERNO COMPLETO CLR 70", "INTERNO COMPLETO CLR 70H", 
    "Fanali C S 155", "Fanali C S 250", "Fanali C L 299", 
    "Fanali C L 299 H", "Fanali C L 301 H", "Fanali C L 401 H", 
    "Fotocellula", "Regolatore di Carica", 
    "Alimentatore 220VAC/ 12VCC 1A", "Batteria", "Pannello Solare", "Lampade"
]
COLORI_LUCE = ["Bianca", "Rossa", "Verde", "Gialla", "Settorizzata", "Oscurato"]
GPS_OPZIONI = ["NO", "SI"]

st.divider()

# --- 1. DATI ORDINE (FISSI) ---
st.subheader("Dati Ordine")
col_test1, col_test2 = st.columns(2)
with col_test1:
    cliente = st.text_input("Cliente", placeholder="Inserire cliente...", key="fisso_cliente").upper()
with col_test2:
    commessa = st.text_input("COMMESSA N°", placeholder="Inserire commessa...", key="fisso_commessa").upper()

col_test3, col_test4 = st.columns(2)
with col_test3:
    consegna_date = st.date_input("SPEDIRE ENTRO IL", value=date.today(), format="DD/MM/YYYY", key="fisso_consegna")
    consegna_str = formatta_data_it(consegna_date)
with col_test4:
    dest_scelto = st.radio("DESTINAZIONE MERCE", ["CLIENTE", "RESINEX"], horizontal=True, key="fisso_dest")

st.divider()

# --- 2. AGGIUNGI ARTICOLO ---
st.subheader("Aggiungi articolo")
suffix = st.session_state.form_reset

tipo_int = st.radio("Lavorazione per", ["VENDITA", "RIPARAZIONE"], horizontal=True, key=f"tipo_{suffix}")
prod_scelto = st.selectbox("Tipologia Articolo", PRODOTTI, key=f"prod_{suffix}")

d_art = {"lamp": "", "port": "", "colore": "", "gps": "", "note": ""}
qta = 1

if prod_scelto != "Selezionare...":
    is_fanaleria = "Fanali" in prod_scelto or "INTERNO" in prod_scelto
    if is_fanaleria:
        ca, cb, cc, cd = st.columns(4)
        with ca: d_art["colore"] = st.selectbox("Colore Luce", COLORI_LUCE, key=f"col_{suffix}")
        with cb: qta = st.number_input("Q.TA'", min_value=1, step=1, key=f"qta_{suffix}")
        with cc: d_art["port"] = st.text_input("Portata MN", placeholder=" ", key=f"port_{suffix}")
        with cd: d_art["gps"] = st.radio("Sincro GPS", GPS_OPZIONI, horizontal=True, key=f"gps_{suffix}")
        d_art["lamp"] = st.text_area("Caratteristica di lampeggio", placeholder=" ", key=f"lamp_{suffix}")
        d_art["note"] = st.text_area("NOTE TECNICHE SPECIFICHE", placeholder=" ", key=f"note_{suffix}")
    else:
        qta = st.number_input("Q.TA'", min_value=1, step=1, key=f"qta_gen_{suffix}")
        d_art["note"] = st.text_area("NOTE TECNICHE SPECIFICHE", placeholder=" ", key=f"note_gen_{suffix}")
        d_art["gps"] = "" 

    if st.button("➕ AGGIUNGI ARTICOLO"):
        if cliente and commessa:
            if is_fanaleria and (not d_art["port"] or not d_art["lamp"]):
                st.warning("⚠️ Portata o Lampeggio sono vuoti.")
                st.info("Spunta la casella sotto per confermare l'inserimento incompleto.")
            else:
                st.session_state.lista_articoli.append({
                    "tipo": tipo_int, "prod": prod_scelto, "qta": qta,
                    "lamp": d_art["lamp"].upper(), "port": d_art["port"].upper(),
                    "colore": d_art["colore"].upper(), "gps": d_art["gps"],
                    "note": d_art["note"].upper()
                })
                st.session_state.form_reset += 1 
                st.rerun()
        else:
            st.error("Inserisci Cliente e Commessa!")

    if is_fanaleria and (not d_art["port"] or not d_art["lamp"]):
        if st.checkbox("Voglio inserire l'articolo anche se incompleto", key=f"chk_{suffix}"):
            if st.button("🚀 CONFERMA AGGIUNGA"):
                st.session_state.lista_articoli.append({
                    "tipo": tipo_int, "prod": prod_scelto, "qta": qta,
                    "lamp": d_art["lamp"].upper(), "port": d_art["port"].upper(),
                    "colore": d_art["colore"].upper(), "gps": d_art["gps"],
                    "note": d_art["note"].upper()
                })
                st.session_state.form_reset += 1 
                st.rerun()

# --- 3. RIEPILOGO E SALVATAGGIO ---
if st.session_state.lista_articoli:
    st.divider()
    st.subheader(f"📋 Articoli inseriti ({len(st.session_state.lista_articoli)})")
    st.dataframe(pd.DataFrame(st.session_state.lista_articoli), use_container_width=True)

    if st.button("🚀 GENERA FILE EXCEL"):
        try:
            wb = load_workbook(MODELLO_EXCEL)
            ws = wb.active 
            ws["D4"] = cliente
            ws["D5"] = commessa
            ws["D42"] = consegna_str
            ws["D44"] = dest_scelto
            ws["E46"] = f"DATA EMISSIONE {formatta_data_it(datetime.now())}"
            
            for i, art in enumerate(st.session_state.lista_articoli[:6]):
                r_label = 7 if i == 0 else (20 if i == 1 else (31 if i == 2 else (52 if i == 3 else (65 if i == 4 else 76))))
                r_base = r_label + 3
                ws[f"D{r_label}"] = art["tipo"]; ws[f"D{r_base}"] = art["prod"]
                ws[f"D{r_base+1}"] = art["qta"]; ws[f"D{r_base+2}"] = art["lamp"]
                ws[f"D{r_base+3}"] = art["port"]; ws[f"D{r_base+4}"] = art["colore"]
                ws[f"D{r_base+5}"] = art["gps"]; ws[f"D{r_base+6}"] = art["note"] 
            
            nome_file = f"{OUTPUT_FOLDER}/Modulo_{commessa}.xlsx"
            wb.save(nome_file)
            st.success(f"✅ Excel generato")
            with open(nome_file, "rb") as f:
                st.download_button("📥 SCARICA", f, file_name=f"Modulo_{commessa}.xlsx")
        except Exception as e:
            st.error(f"Errore: {e}")

    if st.button("🗑️ CANCELLA TUTTA LA LISTA"):
        st.session_state.lista_articoli = []; st.session_state.form_reset = 0; st.rerun()
