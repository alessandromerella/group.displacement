import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta, date
import holidays
import io
import base64
import time
import re
import requests
import json
import os

st.set_page_config(page_title="Hotel Groups Displacement Analyzer v0.9.0beta2", layout="wide")

COLOR_PALETTE = {
    "primary": "#D8C0B7",
    "secondary": "#8CA68C",
    "text": "#5E5E5E",
    "background": "#F8F6F4",
    "accent": "#B6805C",
    "positive": "#8CA68C",
    "negative": "#D8837F"
}

def load_changelog():
    try:
        with open("changelog.md", "r") as f:
            return f.read()
    except FileNotFoundError:
        return """# Changelog Hotel Group Displacement Analyzer

## v0.9.0beta2 (Attuale)
- Migliorata l'interfaccia utente con tabelle a colori differenziate
- Corretto il bug nell'identificazione dei file Excel
- Aggiunta visualizzazione del changelog al primo avvio
- Corretta la visualizzazione delle tabelle di inserimento dati
- Migliorato il processo di autenticazione

## v0.9.0beta1
- **Nuova funzionalità**: Ragionamento Esteso per analisi avanzata
  - Analisi automatica di scenari multipli con variazioni ADR (-10%, -5%, base, +5%, +10%)
  - Identificazione e suggerimento della tariffa ottimale per massimizzare il profitto
  - Visualizzazione grafica dell'impatto delle variazioni di ADR
  - Analisi dei "shoulder days" e giorni critici (per la modalità import Excel)

## v0.8.0
- **Calendario eventi integrato** con dati caricati da server remoto (JSON)
- Supporto a eventi per diverse città
- Avvisi automatici per eventi che coincidono con il periodo di analisi

## v0.7.0
- **Implementazione della modalità import da file Excel**
- Riconoscimento automatico dei tipi di file

## v0.6.0
- **Nuova funzionalità**: Serie di Gruppi per analizzare gruppi ripetitivi

## v0.5.6 e v0.5.5
- Modalità Wizard per inserimento dati guidato
- Miglioramenti stabilità e performance

## v0.5.0 (Versione iniziale)
- Prima release pubblica
- Sistema di autenticazione e gestione sessioni
- Analisi di base per displacement gruppi
"""

def authenticate():
    if 'authenticated' in st.session_state and st.session_state['authenticated']:
        if 'login_time' in st.session_state:
            if time.time() - st.session_state['login_time'] > 28800:
                st.session_state['authenticated'] = False
                st.warning("Sessione scaduta. Effettua nuovamente il login.")
                return False
        return True
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        try:
            logo_url = "https://www.revguardian.altervista.org/hgd.logo.png"
            st.image(logo_url, use_column_width=True)
        except:
            st.error("Impossibile caricare il logo")
            
        st.markdown("<p style='text-align: center;'>v0.9.0beta2</p>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Accedi per continuare</p>", unsafe_allow_html=True)
    
    try:
        valid_usernames = st.secrets.credentials.usernames
        valid_passwords = st.secrets.credentials.passwords
    except:
        valid_usernames = ["not_defined"]
        valid_passwords = ["v2025"]
        st.warning("Utilizzo credenziali di sviluppo. Non funzionanti in produzione")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")
        
        if submit:
            if username in valid_usernames:
                idx = valid_usernames.index(username)
                if idx < len(valid_passwords) and password == valid_passwords[idx]:
                    st.session_state['authenticated'] = True
                    st.session_state['username'] = username
                    st.session_state['login_time'] = time.time()
                    st.success("Login effettuato con successo!")
                    st.rerun()
                else:
                    st.error("Password non valida")
            else:
                st.error("Username non riconosciuto")
    
    return False

if not authenticate():
    st.stop()

if 'has_seen_changelog' not in st.session_state:
    st.markdown("## 🆕 Novità in Hotel Group Displacement Analyzer")
    st.markdown(load_changelog())
    
    if st.button("Ho capito", type="primary"):
        st.session_state['has_seen_changelog'] = True
        st.rerun()
    
    st.stop()

st.sidebar.info(f"Accesso effettuato come: {st.session_state['username']}")
if st.sidebar.button("Logout"):
    for key in ['authenticated', 'username', 'login_time', 'analysis_phase', 'wizard_step', 'forecast_method', 
               'pickup_factor', 'pickup_percentage', 'pickup_value', 'series_data', 'current_passage', 
               'series_complete', 'raw_excel_data', 'available_dates', 'analyzed_data', 'selected_start_date', 
               'selected_end_date', 'events_data_cache', 'events_data_updated', 'enable_extended_reasoning',
               'has_seen_changelog']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

st.markdown(f"""
<style>
    .main .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
        background-color: {COLOR_PALETTE["background"]};
    }}
    h1, h2, h3 {{
        font-family: 'Inter', sans-serif;
        color: {COLOR_PALETTE["text"]};
        font-weight: 500;
    }}
    p, li, div {{
        font-family: 'Inter', sans-serif;
        color: {COLOR_PALETTE["text"]};
    }}
    .stButton button {{
        background-color: {COLOR_PALETTE["primary"]};
        color: white;
        border: none;
        font-family: 'Inter', sans-serif;
    }}
    .stButton button:hover {{
        background-color: {COLOR_PALETTE["accent"]};
    }}
    .stDataFrame {{
        font-family: 'Inter', sans-serif;
    }}
    .stDataFrame table {{
        border-collapse: collapse;
    }}
    .stDataFrame td, .stDataFrame th {{
        border: 1px solid #ddd;
    }}
    
    .corrente-container {{
        border: 3px solid {COLOR_PALETTE["secondary"]} !important;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 20px;
    }}
    
    .precedente-container {{
        border: 3px solid {COLOR_PALETTE["accent"]} !important;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 20px;
    }}
    
    .corrente-header {{
        color: {COLOR_PALETTE["secondary"]};
        font-weight: bold;
        margin-bottom: 10px;
    }}
    
    .precedente-header {{
        color: {COLOR_PALETTE["accent"]};
        font-weight: bold;
        margin-bottom: 10px;
    }}
</style>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)

italian_holidays = holidays.IT()

def is_holiday(day):
    return day in italian_holidays

def same_day_last_year(current_date):
    days_to_subtract = 365
    if (current_date.year % 4 == 0 and current_date.year % 100 != 0) or (current_date.year % 400 == 0):
        days_to_subtract = 366
    
    last_year_date = current_date - timedelta(days=days_to_subtract)
    
    if last_year_date.weekday() != current_date.weekday():
        days_diff = (current_date.weekday() - last_year_date.weekday()) % 7
        last_year_date = last_year_date + timedelta(days=days_diff)
    
    return last_year_date

def get_csv_download_link(df, filename, text):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}.csv">{text}</a>'
    return href

def generate_auth_email(group_name, total_revenue, dates, rooms, adr):
    email_template = f"""
    Oggetto: Richiesta autorizzazione gruppo {group_name} - Valore €{total_revenue:,.2f}
    
    Gentile Revenue Manager,
    
    Richiedo autorizzazione per l'offerta al gruppo "{group_name}" che supera la soglia di €35.000.
    
    Dettagli della richiesta:
    - Date soggiorno: dal {dates[0].strftime('%d/%m/%Y')} al {dates[-1].strftime('%d/%m/%Y')}
    - Numero camere: {rooms} ROH
    - ADR: €{adr:.2f}
    - Valore totale: €{total_revenue:,.2f}
    
    Analisi displacement allegata.
    
    In attesa di riscontro.
    
    Cordiali saluti,
    {st.session_state['username']}
    """
    return email_template

def generate_series_auth_email(group_name, total_series_revenue, total_series_impact, num_passages, series_summary):
    email_template = f"""
    Oggetto: Richiesta autorizzazione SERIE gruppo {group_name} - Valore Totale €{total_series_revenue:,.2f}
    
    Gentile Revenue Manager,
    
    Richiedo autorizzazione per l'offerta alla SERIE del gruppo "{group_name}" che supera la soglia di €35.000.
    
    Dettagli della richiesta:
    - Numero passaggi: {num_passages}
    - Valore totale serie: €{total_series_revenue:,.2f}
    - Impatto totale: €{total_series_impact:,.2f}
    
    Dettaglio passaggi:
    {series_summary.to_string(index=False)}
    
    Analisi displacement allegata per ogni passaggio.
    
    In attesa di riscontro.
    
    Cordiali saluti,
    {st.session_state['username']}
    """
    return email_template

def load_events_from_json_url(city, url="https://www.revguardian.altervista.org/eventi.json"):
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            all_events = json.loads(response.text)
            if city in all_events:
                events_data = all_events[city]
                return pd.DataFrame(events_data)
            else:
                st.warning(f"Nessun evento trovato per {city}")
                return pd.DataFrame()
        else:
            st.error(f"Errore nel caricamento degli eventi: {response.status_code}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Errore nella richiesta degli eventi: {e}")
        return pd.DataFrame()

def get_overlapping_events(events_df, start_date, end_date):
    if events_df.empty:
        return pd.DataFrame()
    
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    overlapping = events_df[
        ((events_df["data_inizio"] >= start_date) & (events_df["data_inizio"] <= end_date)) | 
        ((events_df["data_fine"] >= start_date) & (events_df["data_fine"] <= end_date)) |
        ((events_df["data_inizio"] <= start_date) & (events_df["data_fine"] >= end_date))
    ]
    
    return overlapping

def identify_excel_file_type(df):
    try:
        found_filter_row = False
        filter_text = ""
        
        for idx, row in df.iterrows():
            row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
            if 'Filtri applicati:' in row_str:
                filter_text = row_str
                found_filter_row = True
                break
        
        if not found_filter_row:
            return "UNKNOWN", None, None
        
        year_match = re.search(r'(\d{4})\s+\(S_Esercizio\)', filter_text)
        month_match = re.search(r'(\w+)\s+(\d{4})\s+\(S_Anno\s+Mese\)', filter_text)
        
        year = year_match.group(1) if year_match else None
        month_year = month_match.group(0).split('(')[0].strip() if month_match else None
        
        if "Descrizione Mercato TOB non è Gruppi" in filter_text:
            return "IDV", year, month_year
        elif "Descrizione Mercato TOB è Gruppi" in filter_text:
            return "GRP", year, month_year
        else:
            return "UNKNOWN", year, month_year
    
    except Exception as e:
        st.error(f"Errore nell'identificazione del tipo di file: {e}")
        return "UNKNOWN", None, None

def process_excel_import(uploaded_files):
    if not uploaded_files:
        return None, None, None, None
    
    idv_cy_data = None
    idv_ly_data = None
    grp_otb_data = None
    grp_opz_data = None
    
    current_year = datetime.now().year
    
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file)
            file_type, year, month_year = identify_excel_file_type(df)
            
            data_rows = df[df.iloc[:, 0].str.contains('Giorno', na=False)].index
            if len(data_rows) == 0:
                st.error(f"Formato del file {uploaded_file.name} non riconosciuto: nessuna colonna 'Giorno' trovata")
                continue
                
            data_rows = data_rows[0]
            headers = df.iloc[data_rows].values.tolist()
            
            data_df = df.iloc[data_rows+1:].reset_index(drop=True)
            data_df.columns = headers[:13]
            
            data_df = data_df[~data_df['Giorno'].isna()]
            data_df = data_df[~data_df['Giorno'].str.contains('Filtri applicati:', na=False)]
            
            numeric_cols = ['Room nights', 'Bed nights', 'ADR Cam', 'ADR Bed', 'Room Revenue', 'RevPar']
            for col in numeric_cols:
                if col in data_df.columns:
                    data_df[col] = pd.to_numeric(data_df[col], errors='coerce')
            
            data_df['Giorno'] = pd.to_datetime(data_df['Giorno'], errors='coerce')
            
            if file_type == "IDV":
                if str(current_year) in year or str(current_year) in str(month_year):
                    idv_cy_data = data_df
                    st.success(f"File IDV Anno Corrente riconosciuto: {uploaded_file.name}")
                else:
                    idv_ly_data = data_df
                    st.success(f"File IDV Anno Precedente riconosciuto: {uploaded_file.name}")
            elif file_type == "GRP":
                if data_df['Room nights'].sum() > 0:
                    grp_otb_data = data_df
                    st.success(f"File Gruppi Confermati riconosciuto: {uploaded_file.name}")
                else:
                    grp_opz_data = data_df
                    st.success(f"File Gruppi Opzionati riconosciuto: {uploaded_file.name}")
        
        except Exception as e:
            st.error(f"Errore nell'elaborazione del file {uploaded_file.name}: {e}")
    
    return idv_cy_data, idv_ly_data, grp_otb_data, grp_opz_data

def process_imported_data(idv_cy_data, idv_ly_data, grp_otb_data, grp_opz_data, date_range):
    try:
        result_dates = date_range
        result_df = pd.DataFrame({'data': result_dates})
        
        result_df['giorno'] = result_df['data'].dt.strftime('%a')
        result_df['data_ly'] = result_df['data'].apply(same_day_last_year)
        result_df['giorno_ly'] = result_df['data_ly'].dt.strftime('%a')
        
        idv_cy_data.rename(columns={
            'Giorno': 'data',
            'Room nights': 'otb_ind_rn',
            'ADR Cam': 'otb_ind_adr'
        }, inplace=True)
        
        idv_cy_data['data'] = pd.to_datetime(idv_cy_data['data'])
        idv_cy_filtered = idv_cy_data[idv_cy_data['data'].isin(result_df['data'])]
        
        if not idv_cy_filtered.empty:
            idv_cy_grouped = idv_cy_filtered.groupby('data').agg({
                'otb_ind_rn': 'sum',
                'otb_ind_adr': 'mean'
            }).reset_index()
            
            result_df = pd.merge(result_df, idv_cy_grouped[['data', 'otb_ind_rn', 'otb_ind_adr']], 
                                on='data', how='left')
        else:
            result_df['otb_ind_rn'] = 0
            result_df['otb_ind_adr'] = 0
        
        idv_ly_data.rename(columns={
            'Giorno': 'data',
            'Room nights': 'ly_ind_rn',
            'ADR Cam': 'ly_ind_adr'
        }, inplace=True)
        
        idv_ly_data['data'] = pd.to_datetime(idv_ly_data['data'])
        
        ly_dates = [same_day_last_year(d) for d in result_df['data']]
        idv_ly_filtered = idv_ly_data[idv_ly_data['data'].isin(ly_dates)]
        
        if not idv_ly_filtered.empty:
            idv_ly_filtered['data_ly'] = idv_ly_filtered['data']
            
            idv_ly_grouped = idv_ly_filtered.groupby('data_ly').agg({
                'ly_ind_rn': 'sum',
                'ly_ind_adr': 'mean'
            }).reset_index()
            
            result_df = pd.merge(result_df, idv_ly_grouped[['data_ly', 'ly_ind_rn', 'ly_ind_adr']], 
                                on='data_ly', how='left')
        else:
            result_df['ly_ind_rn'] = 0
            result_df['ly_ind_adr'] = 0
        
        if grp_otb_data is not None and not grp_otb_data.empty:
            grp_otb_data.rename(columns={
                'Giorno': 'data',
                'Room nights': 'grp_otb_rn',
                'ADR Cam': 'grp_otb_adr'
            }, inplace=True)
            
            grp_otb_data['data'] = pd.to_datetime(grp_otb_data['data'])
            grp_otb_filtered = grp_otb_data[grp_otb_data['data'].isin(result_df['data'])]
            
            if not grp_otb_filtered.empty:
                grp_otb_grouped = grp_otb_filtered.groupby('data').agg({
                    'grp_otb_rn': 'sum',
                    'grp_otb_adr': 'mean'
                }).reset_index()
                
                result_df = pd.merge(result_df, grp_otb_grouped[['data', 'grp_otb_rn', 'grp_otb_adr']], 
                                    on='data', how='left')
            else:
                result_df['grp_otb_rn'] = 0
                result_df['grp_otb_adr'] = 0
        else:
            result_df['grp_otb_rn'] = 0
            result_df['grp_otb_adr'] = 0
        
        if grp_opz_data is not None and not grp_opz_data.empty:
            grp_opz_data.rename(columns={
                'Giorno': 'data',
                'Room nights': 'grp_opz_rn',
                'ADR Cam': 'grp_opz_adr'
            }, inplace=True)
            
            grp_opz_data['data'] = pd.to_datetime(grp_opz_data['data'])
            grp_opz_filtered = grp_opz_data[grp_opz_data['data'].isin(result_df['data'])]
            
            if not grp_opz_filtered.empty:
                grp_opz_grouped = grp_opz_filtered.groupby('data').agg({
                    'grp_opz_rn': 'sum',
                    'grp_opz_adr': 'mean'
                }).reset_index()
                
                result_df = pd.merge(result_df, grp_opz_grouped[['data', 'grp_opz_rn', 'grp_opz_adr']], 
                                    on='data', how='left')
            else:
                result_df['grp_opz_rn'] = 0
                result_df['grp_opz_adr'] = 0
        else:
            result_df['grp_opz_rn'] = 0
            result_df['grp_opz_adr'] = 0
        
        result_df = result_df.fillna(0)
        
        result_df['fcst_ind_rn'] = result_df['ly_ind_rn'] * 1.1
        result_df['fcst_ind_adr'] = result_df['otb_ind_adr']
        
        result_df['otb_ind_rev'] = result_df['otb_ind_rn'] * result_df['otb_ind_adr']
        result_df['ly_ind_rev'] = result_df['ly_ind_rn'] * result_df['ly_ind_adr']
        result_df['grp_otb_rev'] = result_df['grp_otb_rn'] * result_df['grp_otb_adr']
        result_df['grp_opz_rev'] = result_df['grp_opz_rn'] * result_df['grp_opz_adr']
        result_df['fcst_ind_rev'] = result_df['fcst_ind_rn'] * result_df['fcst_ind_adr']
        
        result_df['finale_rn'] = result_df['fcst_ind_rn'] + result_df['grp_otb_rn']
        result_df['finale_opz_rn'] = result_df['finale_rn'] + result_df['grp_opz_rn']
        
        result_df['finale_rev'] = result_df['fcst_ind_rev'] + result_df['grp_otb_rev']
        result_df['finale_adr'] = np.where(result_df['finale_rn'] > 0,
                                       result_df['finale_rev'] / result_df['finale_rn'],
                                       0)
        
        return result_df
        
    except Exception as e:
        st.error(f"Errore nell'elaborazione dei dati importati: {e}")
        return None

def process_uploaded_files(idv_cy_file, idv_ly_file, grp_otb_file, grp_opz_file, date_range, date_column_name, rn_column_name, adr_column_name):
    try:
        if idv_cy_file is not None:
            idv_cy_data = pd.read_excel(idv_cy_file)
        else:
            st.error("File dati IDV anno corrente mancante")
            return None
            
        if idv_ly_file is not None:
            idv_ly_data = pd.read_excel(idv_ly_file)
        else:
            st.error("File dati IDV anno precedente mancante")
            return None
        
        grp_otb_data = pd.read_excel(grp_otb_file) if grp_otb_file is not None else pd.DataFrame()
        grp_opz_data = pd.read_excel(grp_opz_file) if grp_opz_file is not None else pd.DataFrame()
        
        idv_cy_data[date_column_name] = pd.to_datetime(idv_cy_data[date_column_name])
        idv_ly_data[date_column_name] = pd.to_datetime(idv_ly_data[date_column_name])
        
        if not grp_otb_data.empty:
            grp_otb_data[date_column_name] = pd.to_datetime(grp_otb_data[date_column_name])
        
        if not grp_opz_data.empty:
            grp_opz_data[date_column_name] = pd.to_datetime(grp_opz_data[date_column_name])
        
        result_dates = pd.date_range(start=date_range[0], end=date_range[-1])
        result_df = pd.DataFrame({'data': result_dates})
        
        result_df['giorno'] = result_df['data'].dt.strftime('%a')
        result_df['data_ly'] = result_df['data'].apply(same_day_last_year)
        result_df['giorno_ly'] = result_df['data_ly'].dt.strftime('%a')
        
        def filter_and_agg(df, date_col, value_cols, dates):
            filtered = df[df[date_col].isin(dates)]
            if filtered.empty:
                return pd.DataFrame({'data': dates, **{col: [0] * len(dates) for col in value_cols}})
            
            grouped = filtered.groupby(date_col).agg({
                col: 'sum' if col == rn_column_name else 'mean' for col in value_cols
            }).reset_index()
            
            grouped = grouped.rename(columns={date_col: 'data'})
            return grouped
        
        idv_cy_filtered = filter_and_agg(
            idv_cy_data, 
            date_column_name, 
            [rn_column_name, adr_column_name], 
            result_df['data']
        )
        
        idv_ly_filtered = filter_and_agg(
            idv_ly_data, 
            date_column_name, 
            [rn_column_name, adr_column_name], 
            result_df['data_ly']
        )
        
        idv_cy_filtered = idv_cy_filtered.rename(columns={
            rn_column_name: 'otb_ind_rn', 
            adr_column_name: 'otb_ind_adr'
        })
        
        idv_ly_filtered = idv_ly_filtered.rename(columns={
            rn_column_name: 'ly_ind_rn', 
            adr_column_name: 'ly_ind_adr'
        })
        idv_ly_filtered = idv_ly_filtered.rename(columns={'data': 'data_ly'})
        
        result_df = pd.merge(result_df, idv_cy_filtered, on='data', how='left')
        
        result_df = pd.merge(result_df, idv_ly_filtered, on='data_ly', how='left')
        
        if not grp_otb_data.empty:
            grp_otb_filtered = filter_and_agg(
                grp_otb_data, 
                date_column_name, 
                [rn_column_name, adr_column_name], 
                result_df['data']
            )
            grp_otb_filtered = grp_otb_filtered.rename(columns={
                rn_column_name: 'grp_otb_rn', 
                adr_column_name: 'grp_otb_adr'
            })
            result_df = pd.merge(result_df, grp_otb_filtered, on='data', how='left')
        else:
            result_df['grp_otb_rn'] = 0
            result_df['grp_otb_adr'] = 0
        
        if not grp_opz_data.empty:
            grp_opz_filtered = filter_and_agg(
                grp_opz_data, 
                date_column_name, 
                [rn_column_name, adr_column_name], 
                result_df['data']
            )
            grp_opz_filtered = grp_opz_filtered.rename(columns={
                rn_column_name: 'grp_opz_rn', 
                adr_column_name: 'grp_opz_adr'
            })
            result_df = pd.merge(result_df, grp_opz_filtered, on='data', how='left')
        else:
            result_df['grp_opz_rn'] = 0
            result_df['grp_opz_adr'] = 0
        
        result_df = result_df.fillna(0)
        
        result_df['fcst_ind_rn'] = result_df['ly_ind_rn'] * 1.1
        result_df['fcst_ind_adr'] = result_df['otb_ind_adr']
        
        result_df['otb_ind_rev'] = result_df['otb_ind_rn'] * result_df['otb_ind_adr']
        result_df['ly_ind_rev'] = result_df['ly_ind_rn'] * result_df['ly_ind_adr']
        result_df['grp_otb_rev'] = result_df['grp_otb_rn'] * result_df['grp_otb_adr']
        result_df['grp_opz_rev'] = result_df['grp_opz_rn'] * result_df['grp_opz_adr']
        result_df['fcst_ind_rev'] = result_df['fcst_ind_rn'] * result_df['fcst_ind_adr']
        
        result_df['finale_rn'] = result_df['fcst_ind_rn'] + result_df['grp_otb_rn']
        result_df['finale_opz_rn'] = result_df['finale_rn'] + result_df['grp_opz_rn']
        
        result_df['finale_rev'] = result_df['fcst_ind_rev'] + result_df['grp_otb_rev']
        result_df['finale_adr'] = np.where(result_df['finale_rn'] > 0,
                                       result_df['finale_rev'] / result_df['finale_rn'],
                                       0)
        
        return result_df
        
    except Exception as e:
        st.error(f"Errore nell'elaborazione dei file: {e}")
        return None

class ExcelCompatibleDisplacementAnalyzer:
    def __init__(self, hotel_capacity, iva_rate=0.1):
        self.hotel_capacity = hotel_capacity
        self.iva_rate = iva_rate
        self.data = None
        self.group_request = None
        self.decision_params = None
    
    def set_data(self, data_df):
        self.data = data_df
        return self
    
    def set_group_request(self, start_date, end_date, num_rooms, adr_lordo, adr_netto=None, fb_revenue=0, meeting_revenue=0, other_revenue=0):
        if adr_netto is None:
            adr_netto = adr_lordo / (1 + self.iva_rate)
            
        date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        total_days = len(date_range)
        
        daily_fb = fb_revenue / total_days if total_days > 0 else 0
        daily_meeting = meeting_revenue / total_days if total_days > 0 else 0
        daily_other = other_revenue / total_days if total_days > 0 else 0
        total_ancillary = daily_fb + daily_meeting + daily_other
        
        self.group_request = pd.DataFrame({
            'data': date_range,
            'camere_gruppo': num_rooms,
            'adr_gruppo_lordo': adr_lordo,
            'adr_gruppo_netto': adr_netto,
            'revenue_camere_gruppo': num_rooms * adr_netto,
            'revenue_fb_gruppo': daily_fb,
            'revenue_meeting_gruppo': daily_meeting,
            'revenue_other_gruppo': daily_other,
            'revenue_ancillare_gruppo': total_ancillary,
            'revenue_totale_gruppo': (num_rooms * adr_netto) + total_ancillary
        })
        
        return self
    
    def set_decision_parameters(self, params):
        self.decision_params = params
        return self
    
    def analyze(self):
        if self.data is None or self.group_request is None or self.decision_params is None:
            raise ValueError("Dati, richiesta gruppo o parametri decisionali mancanti")
        
        result = pd.merge(self.data, self.group_request, on='data', how='right')
        
        result['finale_rn'] = result['fcst_ind_rn'] + result['grp_otb_rn']
        result['finale_opz_rn'] = result['finale_rn']
        
        result['camere_disponibili'] = self.hotel_capacity - result['finale_rn']
        result['camere_displaced'] = np.where(
            self.hotel_capacity - (result['finale_rn'] + result['camere_gruppo']) > 0,
            0,
            (result['finale_rn'] + result['camere_gruppo']) - self.hotel_capacity
        )
        
        result['camere_gruppo_accettate'] = result['camere_gruppo'] - result['camere_displaced']
        
        result['revenue_displaced'] = result['camere_displaced'] * result['finale_adr']
        
        result['revenue_camere_gruppo_effettivo'] = result['camere_gruppo_accettate'] * result['adr_gruppo_netto']
        
        result['impatto_revenue_camera'] = result['revenue_camere_gruppo_effettivo'] - result['revenue_displaced']
        
        result['impatto_revenue_totale'] = result['impatto_revenue_camera'] + result['revenue_ancillare_gruppo']
        
        result['occupazione_attuale'] = result['finale_rn'] / self.hotel_capacity * 100
        result['occupazione_con_gruppo'] = (result['finale_rn'] + result['camere_gruppo_accettate'] - result['camere_displaced']) / self.hotel_capacity * 100
        
        if result['otb_ind_rn'].sum() > 0:
            avg_adr_cy = np.average(result['otb_ind_adr'], weights=result['otb_ind_rn'])
        else:
            avg_adr_cy = result['otb_ind_adr'].mean() if len(result['otb_ind_adr']) > 0 else 0
            
        if result['ly_ind_rn'].sum() > 0:
            avg_adr_ly = np.average(result['ly_ind_adr'], weights=result['ly_ind_rn'])
        else:
            avg_adr_ly = result['ly_ind_adr'].mean() if len(result['ly_ind_adr']) > 0 else 0
        
        result['avg_adr_cy'] = avg_adr_cy
        result['avg_adr_ly'] = avg_adr_ly
        
        result['extra_vs_ly'] = result['adr_gruppo_netto'] - avg_adr_ly
        
        return result
    
    def get_summary_metrics(self, analysis_df):
        total_displaced_revenue = analysis_df['revenue_displaced'].sum()
        total_group_room_revenue = analysis_df['revenue_camere_gruppo_effettivo'].sum()
        total_group_ancillary = analysis_df['revenue_ancillare_gruppo'].sum()
        total_impact = analysis_df['impatto_revenue_totale'].sum()
        
        total_lordo = (total_group_room_revenue * (1 + self.iva_rate)) + total_group_ancillary
        needs_authorization = total_lordo > 35000
        
        adr_netto = analysis_df['adr_gruppo_netto'].mean()
        adr_lordo = analysis_df['adr_gruppo_lordo'].mean()
        
        avg_adr_cy = analysis_df['avg_adr_cy'].iloc[0] if not analysis_df.empty else 0
        avg_adr_ly = analysis_df['avg_adr_ly'].iloc[0] if not analysis_df.empty else 0
        
        extra_vs_ly = adr_netto - avg_adr_ly
        
        total_group_rooms = analysis_df['camere_gruppo'].sum()
        accepted_rooms = analysis_df['camere_gruppo_accettate'].sum()
        displaced_rooms = analysis_df['camere_displaced'].sum()
        room_profit = (adr_netto * accepted_rooms) - total_displaced_revenue
        
        total_rev_profit = room_profit + total_group_ancillary
        
        avg_occ_current = analysis_df['occupazione_attuale'].mean()
        avg_occ_with_group = analysis_df['occupazione_con_gruppo'].mean()
        
        should_accept = total_rev_profit > 0
        
        return {
            'revenue_displaced': total_displaced_revenue,
            'group_room_revenue': total_group_room_revenue,
            'group_ancillary': total_group_ancillary,
            'total_impact': total_impact,
            'total_lordo': total_lordo,
            'needs_authorization': needs_authorization,
            'should_accept': should_accept,
            'current_adr_lordo': adr_lordo,
            'current_adr_netto': adr_netto,
            'avg_adr_cy': avg_adr_cy,
            'avg_adr_ly': avg_adr_ly,
            'extra_vs_ly': extra_vs_ly,
            'total_group_rooms': total_group_rooms,
            'accepted_rooms': accepted_rooms,
            'displaced_rooms': displaced_rooms,
            'avg_occ_current': avg_occ_current,
            'avg_occ_with_group': avg_occ_with_group,
            'room_profit': room_profit,
            'total_rev_profit': total_rev_profit,
            'profit_per_room': room_profit / accepted_rooms if accepted_rooms > 0 else 0,
        }
    
    def create_visualizations(self, analysis_df, metrics, events_df=None):
        fig = make_subplots(rows=3, cols=1,
                          shared_xaxes=True,
                          vertical_spacing=0.1,
                          subplot_titles=('Occupazione', 'ADR', 'Impatto Revenue'))
        
        fig.add_trace(
            go.Bar(name='OTB', x=analysis_df['data'], y=analysis_df['finale_rn'],
                  marker_color=COLOR_PALETTE["secondary"]),
            row=1, col=1
        )
        
        fig.add_trace(
            go.Bar(name='Gruppo', x=analysis_df['data'], y=analysis_df['camere_gruppo'],
                  marker_color=COLOR_PALETTE["primary"]),
            row=1, col=1
        )
        
        fig.add_trace(
            go.Scatter(name='Capacità Hotel', x=analysis_df['data'], 
                      y=[self.hotel_capacity]*len(analysis_df),
                      mode='lines', line=dict(color=COLOR_PALETTE["accent"], dash='dash')),
            row=1, col=1
        )
        
        fig.add_trace(
            go.Scatter(name='ADR Gruppo Netto', x=analysis_df['data'], y=analysis_df['adr_gruppo_netto'],
                     mode='lines+markers', marker=dict(color=COLOR_PALETTE["primary"])),
            row=2, col=1
        )
        
        fig.add_trace(
            go.Scatter(name='ADR Periodo CY', x=analysis_df['data'], 
                      y=[metrics['avg_adr_cy']]*len(analysis_df),
                      mode='lines', line=dict(color=COLOR_PALETTE["secondary"])),
            row=2, col=1
        )
        
        fig.add_trace(
            go.Scatter(name='ADR Periodo LY', x=analysis_df['data'], 
                      y=[metrics['avg_adr_ly']]*len(analysis_df),
                      mode='lines', line=dict(color=COLOR_PALETTE["secondary"], dash='dot')),
            row=2, col=1
        )
        
        fig.add_trace(
            go.Bar(name='Revenue Perso', x=analysis_df['data'], 
                  y=analysis_df['revenue_displaced'],
                  marker_color=COLOR_PALETTE["negative"]),
            row=3, col=1
        )
        
        fig.add_trace(
            go.Bar(name='Revenue Camere', x=analysis_df['data'], 
                  y=analysis_df['revenue_camere_gruppo_effettivo'],
                  marker_color=COLOR_PALETTE["secondary"]),
            row=3, col=1
        )
        
        fig.add_trace(
            go.Bar(name='Revenue Ancillare', x=analysis_df['data'], 
                  y=analysis_df['revenue_ancillare_gruppo'],
                  marker_color=COLOR_PALETTE["primary"]),
            row=3, col=1
        )
        
        if events_df is not None and not events_df.empty:
            for _, event in events_df.iterrows():
                opacity = {"Alto": 0.3, "Medio": 0.2, "Basso": 0.1}.get(event["impatto"], 0.1)
                color = {"Alto": "rgba(255, 87, 51, {})", "Medio": "rgba(255, 195, 0, {})", "Basso": "rgba(218, 247, 166, {})"}
                event_color = color.get(event["impatto"], "rgba(200, 200, 200, {})").format(opacity)
                
                fig.add_vrect(
                    x0=event["data_inizio"], 
                    x1=event["data_fine"],
                    fillcolor=event_color,
                    layer="below",
                    line_width=0,
                    annotation_text=event["nome"],
                    annotation_position="top left",
                    row=1, col=1
                )
        
        fig.update_layout(
            title_text='Analisi Displacement Gruppo',
            height=800,
            barmode='stack',
            font_family="Inter, sans-serif",
            plot_bgcolor=COLOR_PALETTE["background"],
            paper_bgcolor=COLOR_PALETTE["background"],
            font_color=COLOR_PALETTE["text"]
        )
        
        fig_summary = go.Figure()
        
        summary_data = [
            metrics['revenue_displaced'],
            metrics['group_room_revenue'],
            metrics['group_ancillary'],
            metrics['total_impact']
        ]
        
        summary_labels = [
            'Revenue Perso',
            'Revenue Camere',
            'Revenue Ancillare',
            'Impatto Totale'
        ]
        
        colors = [COLOR_PALETTE["negative"], COLOR_PALETTE["secondary"], COLOR_PALETTE["primary"], 
                 COLOR_PALETTE["positive"] if metrics['total_impact'] >= 0 else COLOR_PALETTE["negative"]]
        
        fig_summary.add_trace(go.Bar(
            x=summary_labels,
            y=summary_data,
            marker_color=colors,
            name='Impatto Revenue'
        ))
        
        fig_summary.update_layout(
            title_text=f'Riepilogo Finanziario',
            font_family="Inter, sans-serif",
            plot_bgcolor=COLOR_PALETTE["background"],
            paper_bgcolor=COLOR_PALETTE["background"],
            font_color=COLOR_PALETTE["text"]
        )
        
        return fig, fig_summary


st.title("Hotel Group Displacement Analyzer v0.9.0beta2")
st.markdown("*Strumento di analisi richieste preventivo gruppi*")

with st.sidebar:
    st.header("Configurazione Hotel")
    hotel_capacity = st.number_input("Capacità hotel (camere)", min_value=1, value=66)
    iva_rate = st.number_input("Aliquota IVA (%)", min_value=0.0, max_value=30.0, value=10.0) / 100
    
    st.header("Eventi & Fiere")
    city = st.selectbox("Città", ["Venezia", "Roma", "Taormina", "Olbia", "Cervinia", "Matera", "Siracusa", "Firenze"])
    
    json_url = "https://www.revguardian.altervista.org/eventi.json"
    
    if 'events_data_cache' not in st.session_state:
        with st.spinner("Caricamento database eventi..."):
            try:
                response = requests.get(json_url, timeout=5)
                if response.status_code == 200:
                    st.session_state['events_data_cache'] = json.loads(response.text)
                    st.session_state['events_data_updated'] = datetime.now()
                    st.success("Database eventi caricato con successo!")
                else:
                    st.error("Impossibile caricare il database eventi")
            except Exception as e:
                st.error(f"Errore nella connessione: {e}")
    else:
        last_update = st.session_state.get('events_data_updated', datetime.now())
        st.info(f"Database eventi aggiornato il: {last_update.strftime('%d/%m/%Y %H:%M')}")
        
        if st.button("Aggiorna database"):
            with st.spinner("Aggiornamento database eventi..."):
                try:
                    response = requests.get(json_url, timeout=5)
                    if response.status_code == 200:
                        st.session_state['events_data_cache'] = json.loads(response.text)
                        st.session_state['events_data_updated'] = datetime.now()
                        st.success("Database eventi aggiornato con successo!")
                    else:
                        st.error("Impossibile aggiornare il database eventi")
                except Exception as e:
                    st.error(f"Errore nell'aggiornamento: {e}")
    
    st.header("Impostazioni")
    enable_wizard = st.toggle("Modalità Wizard (guida passo-passo)", value=False, 
                           help="Attiva la guida passo-passo per l'inserimento dei dati")
    
    st.header("Tipo di Analisi")
    enable_series = st.toggle("Serie di Gruppi", value=False, 
                           help="Attiva per analizzare gruppi che si ripetono nel tempo")
    
    if enable_series:
        num_passages = st.number_input("Numero di passaggi", min_value=2, max_value=12, value=3)
    
    st.header("Fonte dati")
    data_source = st.radio("Seleziona fonte dati", ["Import file Excel", "Inserimento manuale"])

if 'series_data' not in st.session_state and enable_series:
    st.session_state['series_data'] = []
    st.session_state['current_passage'] = 1
    st.session_state['series_complete'] = False

if enable_series and not st.session_state.get('series_complete', False):
    st.header(f"Serie Gruppo - Passaggio {st.session_state['current_passage']} di {num_passages}")

if 'events_data_cache' in st.session_state and city in st.session_state['events_data_cache']:
    city_events_data = st.session_state['events_data_cache'][city]
    events_df = pd.DataFrame(city_events_data)
    
    events_df["data_inizio"] = pd.to_datetime(events_df["data_inizio"])
    events_df["data_fine"] = pd.to_datetime(events_df["data_fine"])
else:
    events_df = pd.DataFrame(columns=["data_inizio", "data_fine", "nome", "descrizione", "impatto"])

analyzed_data = None
start_date = None
end_date = None

if data_source == "Import file Excel":
    with st.expander("Caricamento file", expanded=True):
        st.info("Carica i file Excel esportati da Power BI per effettuare l'analisi")
        
        uploaded_files = st.file_uploader("Carica i file Excel (IDV CY, IDV LY, GRP OTB, GRP OPZ)", 
                                         type=["xlsx", "xls"], accept_multiple_files=True)
        
        if uploaded_files:
            if 'raw_excel_data' not in st.session_state:
                with st.spinner("Analisi file in corso..."):
                    idv_cy_data, idv_ly_data, grp_otb_data, grp_opz_data = process_excel_import(uploaded_files)
                    
                    if idv_cy_data is not None and idv_ly_data is not None:
                        st.session_state['raw_excel_data'] = {
                            'idv_cy': idv_cy_data,
                            'idv_ly': idv_ly_data,
                            'grp_otb': grp_otb_data,
                            'grp_opz': grp_opz_data
                        }
                        
                        available_dates = pd.date_range(
                            start=idv_cy_data['Giorno'].min(),
                            end=idv_cy_data['Giorno'].max()
                        )
                        st.session_state['available_dates'] = available_dates
                        
                        st.success(f"File elaborati con successo! Range date disponibili: {available_dates.min().strftime('%d/%m/%Y')} - {available_dates.max().strftime('%d/%m/%Y')}")
                    else:
                        st.warning("Sono necessari almeno i file IDV anno corrente e anno precedente")
            
            if 'raw_excel_data' in st.session_state:
                st.success(f"File già elaborati. Seleziona il periodo da analizzare.")
                
                available_dates = st.session_state['available_dates']
                
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input(
                        "Data inizio analisi", 
                        value=available_dates.min(),
                        min_value=available_dates.min(),
                        max_value=available_dates.max()
                    )
                
                with col2:
                    end_date = st.date_input(
                        "Data fine analisi", 
                        value=available_dates.min() + timedelta(days=3),
                        min_value=available_dates.min(),
                        max_value=available_dates.max()
                    )
                
                if st.button("Carica dati per il periodo selezionato", type="primary"):
                    with st.spinner("Elaborazione dati in corso..."):
                        start_datetime = pd.to_datetime(start_date)
                        end_datetime = pd.to_datetime(end_date)
                        
                        date_range = pd.date_range(start=start_datetime, end=end_datetime)
                        
                        processed_data = process_imported_data(
                            st.session_state['raw_excel_data']['idv_cy'],
                            st.session_state['raw_excel_data']['idv_ly'],
                            st.session_state['raw_excel_data']['grp_otb'],
                            st.session_state['raw_excel_data']['grp_opz'],
                            date_range
                        )
                        
                        if processed_data is not None:
                            st.session_state['analyzed_data'] = processed_data
                            st.session_state['selected_start_date'] = start_date
                            st.session_state['selected_end_date'] = end_date
                            st.rerun()
            
            if st.button("Cambia file", key="reset_excel_data"):
                if 'raw_excel_data' in st.session_state:
                    del st.session_state['raw_excel_data']
                if 'available_dates' in st.session_state:
                    del st.session_state['available_dates']
                if 'analyzed_data' in st.session_state:
                    del st.session_state['analyzed_data']
                if 'selected_start_date' in st.session_state:
                    del st.session_state['selected_start_date']
                if 'selected_end_date' in st.session_state:
                    del st.session_state['selected_end_date']
                st.rerun()
    
    if 'analyzed_data' in st.session_state:
        analyzed_data = st.session_state['analyzed_data']
        start_date = st.session_state['selected_start_date']
        end_date = st.session_state['selected_end_date']
        
        st.subheader("Dati elaborati")
        tab1, tab2, tab3 = st.tabs(["Room Nights", "ADR", "Revenue"])
        
        with tab1:
            st.dataframe(
                analyzed_data[['data', 'giorno', 'otb_ind_rn', 'ly_ind_rn', 'fcst_ind_rn', 'grp_otb_rn', 'grp_opz_rn', 'finale_rn']],
                column_config={
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                    "giorno": "Giorno",
                    "otb_ind_rn": st.column_config.NumberColumn("OTB IND", format="%d"),
                    "ly_ind_rn": st.column_config.NumberColumn("LY IND", format="%d"),
                    "fcst_ind_rn": st.column_config.NumberColumn("FCST IND", format="%d"),
                    "grp_otb_rn": st.column_config.NumberColumn("GRP OTB", format="%d"),
                    "grp_opz_rn": st.column_config.NumberColumn("GRP OPZ", format="%d"),
                    "finale_rn": st.column_config.NumberColumn("TOTALE", format="%d")
                }
            )
        
        with tab2:
            st.dataframe(
                analyzed_data[['data', 'giorno', 'otb_ind_adr', 'ly_ind_adr', 'fcst_ind_adr', 'grp_otb_adr', 'grp_opz_adr', 'finale_adr']],
                column_config={
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                    "giorno": "Giorno",
                    "otb_ind_adr": st.column_config.NumberColumn("OTB IND", format="€%.2f"),
                    "ly_ind_adr": st.column_config.NumberColumn("LY IND", format="€%.2f"),
                    "fcst_ind_adr": st.column_config.NumberColumn("FCST IND", format="€%.2f"),
                    "grp_otb_adr": st.column_config.NumberColumn("GRP OTB", format="€%.2f"),
                    "grp_opz_adr": st.column_config.NumberColumn("GRP OPZ", format="€%.2f"),
                    "finale_adr": st.column_config.NumberColumn("FINALE", format="€%.2f")
                }
            )
        
        with tab3:
            st.dataframe(
                analyzed_data[['data', 'giorno', 'otb_ind_rev', 'ly_ind_rev', 'fcst_ind_rev', 'grp_otb_rev', 'grp_opz_rev', 'finale_rev']],
                column_config={
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                    "giorno": "Giorno",
                    "otb_ind_rev": st.column_config.NumberColumn("OTB IND", format="€%.2f"),
                    "ly_ind_rev": st.column_config.NumberColumn("LY IND", format="€%.2f"),
                    "fcst_ind_rev": st.column_config.NumberColumn("FCST IND", format="€%.2f"),
                    "grp_otb_rev": st.column_config.NumberColumn("GRP OTB", format="€%.2f"),
                    "grp_opz_rev": st.column_config.NumberColumn("GRP OPZ", format="€%.2f"),
                    "finale_rev": st.column_config.NumberColumn("FINALE", format="€%.2f")
                }
            )
    
    with st.expander("Mappatura campi (avanzato)", expanded=False):
        st.warning("Queste impostazioni sono utilizzate solo se il riconoscimento automatico fallisce")
        date_column_name = st.text_input("Nome colonna data", "Giorno")
        rn_column_name = st.text_input("Nome colonna room nights", "Room nights")
        adr_column_name = st.text_input("Nome colonna ADR", "ADR Cam")
       
elif data_source == "Inserimento manuale":
    st.header("1️⃣ Periodo di Analisi")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inizio analisi", value=datetime.now() + timedelta(days=30))
    with col2:
        end_date = st.date_input("Data fine analisi", value=datetime.now() + timedelta(days=33))
    
    st.header("2️⃣ Dati On The Books e Forecast")
    
    st.info("Inserisci manualmente i dati per il periodo selezionato")
    
    date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
    
    if 'wizard_step' not in st.session_state and enable_wizard:
        st.session_state['wizard_step'] = 1
    
    if 'forecast_method' not in st.session_state:
        st.session_state['forecast_method'] = "Basato su LY"
        st.session_state['pickup_factor'] = 1.1
        st.session_state['pickup_percentage'] = 20
        st.session_state['pickup_value'] = 10
    
    base_data = {
        'data': date_range,
        'giorno': [d.strftime('%a') for d in date_range],
        'data_ly': [same_day_last_year(d) for d in date_range],
        'giorno_ly': [same_day_last_year(d).strftime('%a') for d in date_range],
        'otb_ind_rn': [0] * len(date_range),
        'ly_ind_rn': [0] * len(date_range),
        'grp_otb_rn': [0] * len(date_range),
        'grp_opz_rn': [0] * len(date_range),
        'otb_ind_adr': [0.0] * len(date_range),
        'ly_ind_adr': [0.0] * len(date_range),
        'grp_otb_adr': [0.0] * len(date_range),
        'grp_opz_adr': [0.0] * len(date_range)
    }

    df_base = pd.DataFrame(base_data)
    
    if enable_wizard:
        wizard_steps = [
            "Room Nights - Anno Corrente",
            "ADR - Anno Corrente",
            "Room Nights - Anno Precedente",
            "ADR - Anno Precedente",
            "Parametri Forecast",
            "Completa"
        ]
        
        st.progress(st.session_state['wizard_step'] / len(wizard_steps))
        current_step = wizard_steps[st.session_state['wizard_step']-1]
        st.info(f"**WIZARD - Passo {st.session_state['wizard_step']}/{len(wizard_steps)}**: {current_step}")
        
        def next_step():
            if st.session_state['wizard_step'] < len(wizard_steps):
                st.session_state['wizard_step'] += 1
                st.rerun()
        
        def prev_step():
            if st.session_state['wizard_step'] > 1:
                st.session_state['wizard_step'] -= 1
                st.rerun()
    
    wizard_visible = not enable_wizard or st.session_state.get('wizard_step') == 1
    if wizard_visible:
        st.subheader("Inserimento Room Nights")
        
        st.markdown('<div class="corrente-container">', unsafe_allow_html=True)
        st.markdown('<div class="corrente-header">Room Nights - Anno Corrente</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            st.markdown("👇 **Compila i valori Room Nights correnti**")
        edited_rn_cy = st.data_editor(
            df_base[['data', 'giorno', 'otb_ind_rn', 'grp_otb_rn', 'grp_opz_rn']],
            hide_index=True,
            key="rn_cy",
            column_config={
                "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                "giorno": "Giorno",
                "otb_ind_rn": st.column_config.NumberColumn("OTB IND", min_value=0, format="%d"),
                "grp_otb_rn": st.column_config.NumberColumn("GRP OTB", min_value=0, format="%d"),
                "grp_opz_rn": st.column_config.NumberColumn("GRP OPZ", min_value=0, format="%d")
            },
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            col1, col2 = st.columns([1, 1])
            with col2:
                if st.button("Avanti →", key="next1", type="primary", use_container_width=True):
                    next_step()
    else:
        edited_rn_cy = st.data_editor(
            df_base[['data', 'giorno', 'otb_ind_rn', 'grp_otb_rn', 'grp_opz_rn']],
            hide_index=True,
            key="rn_cy",
            disabled=True,
            use_container_width=True
        )
        st.markdown("", unsafe_allow_html=True)
    
    wizard_visible = not enable_wizard or st.session_state.get('wizard_step') == 2
    if wizard_visible:
        st.subheader("Inserimento ADR")
        
        st.markdown('<div class="corrente-container">', unsafe_allow_html=True)
        st.markdown('<div class="corrente-header">ADR - Anno Corrente</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            st.markdown("👇 **Compila i valori ADR correnti**")
        edited_adr_cy = st.data_editor(
            df_base[['data', 'giorno', 'otb_ind_adr', 'grp_otb_adr', 'grp_opz_adr']],
            hide_index=True,
            key="adr_cy",
            column_config={
                "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                "giorno": "Giorno",
                "otb_ind_adr": st.column_config.NumberColumn("OTB IND", min_value=0, format="€%.2f"),
                "grp_otb_adr": st.column_config.NumberColumn("GRP OTB", min_value=0, format="€%.2f"),
                "grp_opz_adr": st.column_config.NumberColumn("GRP OPZ", min_value=0, format="€%.2f")
            },
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("← Indietro", key="prev2", use_container_width=True):
                    prev_step()
            with col2:
                if st.button("Avanti →", key="next2", type="primary", use_container_width=True):
                    next_step()
    else:
        edited_adr_cy = st.data_editor(
            df_base[['data', 'giorno', 'otb_ind_adr', 'grp_otb_adr', 'grp_opz_adr']],
            hide_index=True,
            key="adr_cy",
            disabled=True,
            use_container_width=True
        )
        st.markdown("", unsafe_allow_html=True)
    
    wizard_visible = not enable_wizard or st.session_state.get('wizard_step') == 3
    if wizard_visible:
        st.subheader("Room Nights - Anno Precedente")
        
        st.markdown('<div class="precedente-container">', unsafe_allow_html=True)
        st.markdown('<div class="precedente-header">Room Nights - Anno Precedente</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            st.markdown("👇 **Compila i valori Room Nights anno precedente**")
        edited_rn_ly = st.data_editor(
            df_base[['data_ly', 'giorno_ly', 'ly_ind_rn']],
            hide_index=True,
            key="rn_ly",
            column_config={
                "data_ly": st.column_config.DateColumn("Data LY", format="DD/MM/YYYY"),
                "giorno_ly": "Giorno LY",
                "ly_ind_rn": st.column_config.NumberColumn("LY IND", min_value=0, format="%d")
            },
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("← Indietro", key="prev3", use_container_width=True):
                    prev_step()
            with col2:
                if st.button("Avanti →", key="next3", type="primary", use_container_width=True):
                    next_step()
    else:
        edited_rn_ly = st.data_editor(
            df_base[['data_ly', 'giorno_ly', 'ly_ind_rn']],
            hide_index=True,
            key="rn_ly",
            disabled=True,
            use_container_width=True
        )
        st.markdown("", unsafe_allow_html=True)
    
    wizard_visible = not enable_wizard or st.session_state.get('wizard_step') == 4
    if wizard_visible:
        st.subheader("ADR - Anno Precedente")
        
        st.markdown('<div class="precedente-container">', unsafe_allow_html=True)
        st.markdown('<div class="precedente-header">ADR - Anno Precedente</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            st.markdown("👇 **Compila i valori ADR anno precedente**")
        edited_adr_ly = st.data_editor(
            df_base[['data_ly', 'giorno_ly', 'ly_ind_adr']],
            hide_index=True,
            key="adr_ly",
            column_config={
                "data_ly": st.column_config.DateColumn("Data LY", format="DD/MM/YYYY"),
                "giorno_ly": "Giorno LY",
                "ly_ind_adr": st.column_config.NumberColumn("LY IND", min_value=0, format="€%.2f")
            },
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        if enable_wizard:
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("← Indietro", key="prev4", use_container_width=True):
                    prev_step()
            with col2:
                if st.button("Avanti →", key="next4", type="primary", use_container_width=True):
                    next_step()
    else:
        edited_adr_ly = st.data_editor(
            df_base[['data_ly', 'giorno_ly', 'ly_ind_adr']],
            hide_index=True,
            key="adr_ly",
            disabled=True,
            use_container_width=True
        )
        st.markdown("", unsafe_allow_html=True)
    
    wizard_visible = not enable_wizard or st.session_state.get('wizard_step') == 5
    if wizard_visible:
        st.subheader("Parametri Forecast")
        with st.container(border=True):
            if enable_wizard:
                st.markdown("👇 **Imposta i parametri per il forecast**")
            col1, col2 = st.columns(2)
            
            with col1:
                forecast_method = st.selectbox("Metodo di Forecast", 
                                             ["Basato su LY", "Percentuale su OTB", "Valore assoluto"],
                                             index=["Basato su LY", "Percentuale su OTB", "Valore assoluto"].index(st.session_state['forecast_method']))
                st.session_state['forecast_method'] = forecast_method
            
            with col2:
                if forecast_method == "Basato su LY":
                    pickup_factor = st.slider("Moltiplicatore LY", 0.5, 2.0, st.session_state['pickup_factor'], 0.1, 
                                          help="Moltiplica i dati LY per questo fattore")
                    st.session_state['pickup_factor'] = pickup_factor
                elif forecast_method == "Percentuale su OTB":
                    pickup_percentage = st.slider("Pickup %", 0, 100, st.session_state['pickup_percentage'], 5, 
                                              help="Aggiunge questa percentuale all'OTB attuale")
                    st.session_state['pickup_percentage'] = pickup_percentage
                else:
                    pickup_value = st.number_input("Camere da aggiungere", 0, 100, st.session_state['pickup_value'],
                                                help="Aggiunge questo numero di camere all'OTB attuale")
                    st.session_state['pickup_value'] = pickup_value
        
        if enable_wizard:
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("← Indietro", key="prev5", use_container_width=True):
                    prev_step()
            with col2:
                if st.button("Completa", key="complete", type="primary", use_container_width=True):
                    next_step()
    elif enable_wizard and st.session_state.get('wizard_step') == 6:
        st.success("✅ Configurazione completata! Procedi con l'analisi dei dati.")
   
    try:
        final_data = edited_rn_cy.copy()
        final_data = pd.merge(final_data, edited_rn_ly[['data_ly', 'ly_ind_rn']], 
                           left_index=True, right_index=True)
       
        final_data = pd.merge(final_data, edited_adr_cy[['otb_ind_adr', 'grp_otb_adr', 'grp_opz_adr']], 
                           left_index=True, right_index=True)
        final_data = pd.merge(final_data, edited_adr_ly[['ly_ind_adr']], 
                           left_index=True, right_index=True)
       
        if st.session_state['forecast_method'] == "Basato su LY":
            final_data['fcst_ind_rn'] = np.ceil(final_data['ly_ind_rn'] * st.session_state['pickup_factor'])
        elif st.session_state['forecast_method'] == "Percentuale su OTB":
            final_data['fcst_ind_rn'] = np.ceil(final_data['otb_ind_rn'] * (1 + st.session_state['pickup_percentage']/100))
        else:
            final_data['fcst_ind_rn'] = final_data['otb_ind_rn'] + st.session_state['pickup_value']
       
        final_data['fcst_ind_adr'] = final_data['otb_ind_adr']
       
        final_data['finale_rn'] = final_data['fcst_ind_rn'] + final_data['grp_otb_rn']
        final_data['finale_opz_rn'] = final_data['fcst_ind_rn'] + final_data['grp_otb_rn'] + final_data['grp_opz_rn']
       
        final_data['otb_ind_rev'] = final_data['otb_ind_rn'] * final_data['otb_ind_adr']
        final_data['ly_ind_rev'] = final_data['ly_ind_rn'] * final_data['ly_ind_adr']
        final_data['grp_otb_rev'] = final_data['grp_otb_rn'] * final_data['grp_otb_adr']
        final_data['grp_opz_rev'] = final_data['grp_opz_rn'] * final_data['grp_opz_adr']
        final_data['fcst_ind_rev'] = final_data['fcst_ind_rn'] * final_data['fcst_ind_adr']
       
        final_data['finale_rev'] = final_data['fcst_ind_rev'] + final_data['grp_otb_rev']
        final_data['finale_adr'] = np.where(final_data['finale_rn'] > 0,
                                        final_data['finale_rev'] / final_data['finale_rn'],
                                        0)
       
        if not enable_wizard or st.session_state.get('wizard_step') == 6:
            st.subheader("Forecast Calcolato")
            tab1, tab2, tab3 = st.tabs(["Room Nights", "ADR", "Revenue"])
           
            with tab1:
                st.dataframe(
                    final_data[['data', 'giorno', 'otb_ind_rn', 'ly_ind_rn', 'fcst_ind_rn', 'grp_otb_rn', 'grp_opz_rn', 'finale_rn']],
                    column_config={
                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                        "giorno": "Giorno",
                        "otb_ind_rn": st.column_config.NumberColumn("OTB IND", format="%d"),
                        "ly_ind_rn": st.column_config.NumberColumn("LY IND", format="%d"),
                        "fcst_ind_rn": st.column_config.NumberColumn("FCST IND", format="%d"),
                        "grp_otb_rn": st.column_config.NumberColumn("GRP OTB", format="%d"),
                        "grp_opz_rn": st.column_config.NumberColumn("GRP OPZ", format="%d"),
                        "finale_rn": st.column_config.NumberColumn("TOTALE", format="%d")
                    }
                )
           
            with tab2:
                st.dataframe(
                    final_data[['data', 'giorno', 'otb_ind_adr', 'ly_ind_adr', 'fcst_ind_adr', 'grp_otb_adr', 'grp_opz_adr', 'finale_adr']],
                    column_config={
                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                        "giorno": "Giorno",
                        "otb_ind_adr": st.column_config.NumberColumn("OTB IND", format="€%.2f"),
                        "ly_ind_adr": st.column_config.NumberColumn("LY IND", format="€%.2f"),
                        "fcst_ind_adr": st.column_config.NumberColumn("FCST IND", format="€%.2f"),
                        "grp_otb_adr": st.column_config.NumberColumn("GRP OTB", format="€%.2f"),
                        "grp_opz_adr": st.column_config.NumberColumn("GRP OPZ", format="€%.2f"),
                        "finale_adr": st.column_config.NumberColumn("FINALE", format="€%.2f")
                    }
                )
           
            with tab3:
                st.dataframe(
                    final_data[['data', 'giorno', 'otb_ind_rev', 'ly_ind_rev', 'fcst_ind_rev', 'grp_otb_rev', 'grp_opz_rev', 'finale_rev']],
                    column_config={
                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                        "giorno": "Giorno",
                        "otb_ind_rev": st.column_config.NumberColumn("OTB IND", format="€%.2f"),
                        "ly_ind_rev": st.column_config.NumberColumn("LY IND", format="€%.2f"),
                        "fcst_ind_rev": st.column_config.NumberColumn("FCST IND", format="€%.2f"),
                        "grp_otb_rev": st.column_config.NumberColumn("GRP OTB", format="€%.2f"),
                        "grp_opz_rev": st.column_config.NumberColumn("GRP OPZ", format="€%.2f"),
                        "finale_rev": st.column_config.NumberColumn("FINALE", format="€%.2f")
                    }
                )
       
        analyzed_data = final_data
       
    except Exception as e:
        if not enable_wizard or st.session_state.get('wizard_step') == 6:
            st.error(f"Errore nel calcolo del forecast: {e}")
        analyzed_data = None

if start_date is None or end_date is None:
    st.header("1️⃣ Periodo di Analisi")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Data inizio analisi", value=datetime.now() + timedelta(days=30))
    with col2:
        end_date = st.date_input("Data fine analisi", value=datetime.now() + timedelta(days=33))
elif data_source == "Import file Excel" and 'analyzed_data' in st.session_state:
    st.header("1️⃣ Periodo di Analisi")
    st.info(f"Periodo di analisi: dal {start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')}")

st.header("3️⃣ Dettagli Richiesta Gruppo")

col1, col2 = st.columns(2)
with col1:
   group_name = st.text_input("Nome Gruppo", "Corporate Meeting")
   
   st.subheader("Camere ROH")
   num_rooms = st.number_input("Numero camere ROH", min_value=1, value=25)
   
   group_arrival = st.date_input("Data di arrivo", value=start_date, key="group_arrival")
   group_departure = st.date_input("Data di partenza", value=end_date, key="group_departure")
   
with col2:
   adr_lordo = st.number_input("ADR proposta (€ lordi)", min_value=0.0, value=900.0)
   adr_netto = adr_lordo / (1 + iva_rate)
   st.info(f"ADR netto: €{adr_netto:.2f}")
   
   fb_revenue = st.number_input("Revenue F&B previsto (€)", min_value=0.0, value=0.0)
   meeting_revenue = st.number_input("Revenue sale riunioni (€)", min_value=0.0, value=0.0)
   other_revenue = st.number_input("Altro revenue ancillare (€)", min_value=0.0, value=0.0)
   
   total_ancillary = fb_revenue + meeting_revenue + other_revenue
   st.info(f"Totale revenue ancillare: €{total_ancillary:.2f}")
   
   date_nights = (group_departure - group_arrival).days
   total_value = (adr_lordo * num_rooms * date_nights) + total_ancillary
   
   if total_value > 35000:
       st.warning(f"⚠️ Valore totale: €{total_value:,.2f} - Richiede autorizzazione (>€35.000)")
   else:
       st.success(f"✅ Valore totale: €{total_value:,.2f}")

if group_arrival is not None and group_departure is not None:
    overlapping_events = get_overlapping_events(events_df, group_arrival, group_departure)
    
    if not overlapping_events.empty:
        st.warning("⚠️ **ATTENZIONE**: Eventi importanti nel periodo selezionato!")
        
        with st.expander("📅 Eventi nel periodo", expanded=True):
            for _, event in overlapping_events.iterrows():
                impact_color = {
                    "Alto": "#FF5733",
                    "Medio": "#FFC300",
                    "Basso": "#DAF7A6"
                }.get(event["impatto"], "#FFFFFF")
                
                st.markdown(f"""
                <div style="padding: 10px; border-radius: 5px; margin-bottom: 10px; background-color: {impact_color}20; border-left: 5px solid {impact_color};">
                    <h4 style="margin:0;">{event['nome']}</h4>
                    <p style="margin:0; font-size: 0.9em;">📆 {event['data_inizio'].strftime('%d/%m/%Y')} - {event['data_fine'].strftime('%d/%m/%Y')}</p>
                    <p style="margin-top: 5px;">{event['descrizione']}</p>
                    <p style="margin:0; font-weight: bold;">Impatto: {event['impatto']}</p>
                </div>
                """, unsafe_allow_html=True)
                
        high_impact_events = overlapping_events[overlapping_events["impatto"] == "Alto"]
        if not high_impact_events.empty:
            suggested_adr_increase = 15
            suggested_adr = adr_lordo * (1 + suggested_adr_increase/100)
            
            st.info(f"""
            💡 **Suggerimento Revenue**: Il periodo selezionato contiene eventi ad alto impatto. 
            La domanda potrebbe essere significativamente più alta del forecast basato sui dati storici.
            
            Considerando l'evento, valuta un ADR di €{suggested_adr:.2f} (+{suggested_adr_increase}%)
            """)

if start_date is not None and group_arrival is not None and group_departure is not None:
    date_options = pd.date_range(start=group_arrival, end=group_departure - timedelta(days=1))
    formatted_date_options = [f"{d.strftime('%a')} {d.strftime('%d/%m/%Y')}" for d in date_options]
    date_dict = dict(zip(formatted_date_options, date_options))

    selected_formatted_dates = st.multiselect(
        "Seleziona date da includere nell'analisi (lascia vuoto per tutte)",
        options=formatted_date_options,
        default=formatted_date_options
    )

    dates_for_analysis = [date_dict[d] for d in selected_formatted_dates] if selected_formatted_dates else date_options

st.header("4️⃣ Analisi Displacement")

if analyzed_data is None:
    st.error("Nessun dato disponibile per l'analisi. Assicurati di caricare i file necessari o di inserire i dati manualmente.")
else:
    if 'analysis_phase' not in st.session_state:
        st.session_state['analysis_phase'] = 'start'
    
    if st.session_state['analysis_phase'] == 'start':
        if st.button("Esegui Analisi", type="primary", use_container_width=True):
            st.session_state['analysis_phase'] = 'verify'
            st.rerun()
    
    elif st.session_state['analysis_phase'] == 'verify':
        with st.expander("Verifica parametri analisi", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"""
                **Dettagli Gruppo**
                - Nome: {group_name}
                - Periodo: {group_arrival.strftime('%d/%m/%Y')} - {group_departure.strftime('%d/%m/%Y')} ({date_nights} notti)
                - Camere: {num_rooms} ROH
                - Giorni analizzati: {len(dates_for_analysis)} di {len(date_options)}
                """)
            
            with col2:
                st.info(f"""
                **Dettagli Economici**
                - ADR lordo: €{adr_lordo:.2f}
                - ADR netto: €{adr_netto:.2f}
                - Revenue ancillare: €{total_ancillary:.2f}
                - Valore totale: €{total_value:.2f}
                """)
        
        enable_extended_reasoning = st.checkbox("Attiva Ragionamento Esteso", 
                                            help="Esegue un'analisi più approfondita con scenari multipli di ADR e suggerimenti ottimali")
        
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Torna indietro", key="back_to_start"):
                st.session_state['analysis_phase'] = 'start'
                st.rerun()
        with col2:
            if st.button("Conferma e Procedi", key="proceed_to_confirm"):
                st.session_state['enable_extended_reasoning'] = enable_extended_reasoning
                st.session_state['analysis_phase'] = 'confirm'
                st.rerun()
    
    elif st.session_state['analysis_phase'] == 'confirm':
        st.success("✅ Parametri confermati! Clicca 'Conferma Analisi' per procedere.")
        
        if st.session_state.get('enable_extended_reasoning', False):
            st.info("🧠 Ragionamento Esteso attivato: Verranno analizzati scenari multipli con variazioni di ADR")
        
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("Modifica parametri", key="back_to_verify"):
                st.session_state['analysis_phase'] = 'verify'
                st.rerun()
        with col2:
            confirm = st.button("Conferma Analisi", type="primary", use_container_width=True)
            if confirm:
                st.session_state['analysis_phase'] = 'analysis'
                st.rerun()
    
    elif st.session_state['analysis_phase'] == 'analysis':
        with st.spinner("Elaborazione in corso..."):
            analyzer = ExcelCompatibleDisplacementAnalyzer(hotel_capacity=hotel_capacity, iva_rate=iva_rate)
            
            analyzer.set_data(analyzed_data)
            
            decision_parameters = {
                'min_adr_perc_cy': 100,
                'min_adr_perc_ly': 100,
                'ancillary_weight': 1.0,
                'occ_threshold_low': 30,
                'occ_threshold_high': 80,
                'adr_flexibility_low': 0.3,
                'adr_flexibility_high': 0.1,
            }
       
            analyzer.set_decision_parameters(decision_parameters)
       
            analyzer.set_group_request(
               start_date=group_arrival,
               end_date=group_departure,
               num_rooms=num_rooms,
               adr_lordo=adr_lordo,
               adr_netto=adr_netto,
               fb_revenue=fb_revenue,
               meeting_revenue=meeting_revenue,
               other_revenue=other_revenue
            )
       
            result_df = analyzer.analyze()
           
            if dates_for_analysis and len(dates_for_analysis) < len(date_options):
               result_df = result_df[result_df['data'].isin(dates_for_analysis)]
    
            metrics = analyzer.get_summary_metrics(result_df)
            
            if not overlapping_events.empty:
                detail_fig, summary_fig = analyzer.create_visualizations(result_df, metrics, overlapping_events)
            else:
                detail_fig, summary_fig = analyzer.create_visualizations(result_df, metrics)
            
            if st.session_state.get('enable_extended_reasoning', False):
                adr_variations = [-10, -5, 0, 5, 10]
                scenario_results = []
                
                for variation in adr_variations:
                    new_adr = adr_lordo * (1 + variation/100)
                    new_adr_netto = new_adr / (1 + iva_rate)
                    
                    analyzer_scenario = ExcelCompatibleDisplacementAnalyzer(hotel_capacity=hotel_capacity, iva_rate=iva_rate)
                    analyzer_scenario.set_data(analyzed_data)
                    analyzer_scenario.set_decision_parameters(decision_parameters)
                    analyzer_scenario.set_group_request(
                        start_date=group_arrival,
                        end_date=group_departure,
                        num_rooms=num_rooms,
                        adr_lordo=new_adr,
                        fb_revenue=fb_revenue,
                        meeting_revenue=meeting_revenue,
                        other_revenue=other_revenue
                    )
                    
                    result_scenario = analyzer_scenario.analyze()
                    if dates_for_analysis and len(dates_for_analysis) < len(date_options):
                        result_scenario = result_scenario[result_scenario['data'].isin(dates_for_analysis)]
                    
                    metrics_scenario = analyzer_scenario.get_summary_metrics(result_scenario)
                    
                    scenario_results.append({
                        "variation": variation,
                        "adr_lordo": new_adr,
                        "adr_netto": new_adr_netto,
                        "total_impact": metrics_scenario['total_impact'],
                        "room_profit": metrics_scenario['room_profit'],
                        "total_rev_profit": metrics_scenario['total_rev_profit'],
                        "displaced_rooms": metrics_scenario['displaced_rooms'],
                        "should_accept": metrics_scenario['should_accept'],
                        "total_lordo": metrics_scenario['total_lordo']
                    })
                
                scenarios_df = pd.DataFrame(scenario_results)
                
                optimal_scenario = max(scenario_results, key=lambda x: x['total_rev_profit'])
                
                extended_analysis_results = {
                    'scenarios_df': scenarios_df,
                    'optimal_scenario': optimal_scenario
                }
                
                if data_source == "Import file Excel" and 'raw_excel_data' in st.session_state:
                    result_df['criticità'] = pd.cut(
                        result_df['camere_displaced'],
                        bins=[-1, 0, 5, 10, float('inf')],
                        labels=['Nessuna', 'Bassa', 'Media', 'Alta']
                    )
                    
                    extended_analysis_results['critical_days'] = result_df[['data', 'giorno', 'finale_rn', 'camere_gruppo', 'camere_displaced', 'criticità']]
           
        st.subheader("Riepilogo Decisione")
           
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("TOT. LORDO", f"€{metrics['total_lordo']:,.2f}")
        with col2:
            st.metric("TOT. NETTO", f"€{metrics['group_room_revenue'] + metrics['group_ancillary']:,.2f}")
        with col3:
            st.metric("REV DSPL", f"€{metrics['revenue_displaced']:,.2f}")
        with col4:
            st.metric("DIFF", f"€{metrics['total_impact']:,.2f}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(detail_fig, use_container_width=True)
        with col2:
            st.plotly_chart(summary_fig, use_container_width=True)
               
            st.subheader("Riepilogo")
               
            financial_df = pd.DataFrame({
                'Voce': ['TOT. LORDO', 'TOT. NETTO', 'Offerta', 'ADR netto', 'Ancillary', 'Room Profit GROSS', 
                       'Room Profit NET', 'Extra IND LY', 'Extra per room IND LY', 'Extra TY', 'Total Rev Profit'],
                'Valore': [
                    f"€{metrics['total_lordo']:,.2f}",
                    f"€{metrics['group_room_revenue'] + metrics['group_ancillary']:,.2f}",
                    f"€{adr_lordo:.2f}",
                    f"€{adr_netto:.2f}",
                    f"€{metrics['group_ancillary']:,.2f}",
                    f"€{(metrics['group_room_revenue'] * (1 + iva_rate)):,.2f}",
                    f"€{metrics['room_profit']:,.2f}",
                    f"€{(metrics['extra_vs_ly'] * metrics['accepted_rooms']):,.2f}",
                    f"€{metrics['extra_vs_ly']:,.2f}",
                    f"€{metrics['room_profit']:,.2f}",
                    f"€{metrics['total_rev_profit']:,.2f}"
                ]
            })
               
            st.table(financial_df)
        
        if st.session_state.get('enable_extended_reasoning', False) and 'extended_analysis_results' in locals():
            st.header("🧠 Ragionamento Esteso")
            
            st.subheader("Confronto Scenari di ADR")
            st.dataframe(
                extended_analysis_results['scenarios_df'],
                column_config={
                    "variation": st.column_config.NumberColumn("Variazione %", format="%+d%%"),
                    "adr_lordo": st.column_config.NumberColumn("ADR Lordo", format="€%.2f"),
                    "adr_netto": st.column_config.NumberColumn("ADR Netto", format="€%.2f"),
                    "total_impact": st.column_config.NumberColumn("Impatto Rev", format="€%.2f"),
                    "room_profit": st.column_config.NumberColumn("Profitto Camere", format="€%.2f"),
                    "total_rev_profit": st.column_config.NumberColumn("Profitto Totale", format="€%.2f"),
                    "displaced_rooms": st.column_config.NumberColumn("Camere Displaced", format="%d"),
                    "should_accept": st.column_config.CheckboxColumn("Da Accettare")
                }
            )
            
            scenarios_df = extended_analysis_results['scenarios_df']
            fig_scenarios = px.line(
                scenarios_df, 
                x="variation", 
                y=["total_rev_profit", "room_profit"],
                markers=True,
                labels={
                    "variation": "Variazione ADR (%)",
                    "value": "Profitto (€)",
                    "variable": "Tipo"
                },
                title="Impatto delle variazioni di ADR sul profitto",
                color_discrete_map={
                    "total_rev_profit": COLOR_PALETTE["positive"],
                    "room_profit": COLOR_PALETTE["secondary"]
                }
            )
            
            fig_scenarios.update_layout(
                font_family="Inter, sans-serif",
                plot_bgcolor=COLOR_PALETTE["background"],
                paper_bgcolor=COLOR_PALETTE["background"],
                font_color=COLOR_PALETTE["text"]
            )
            
            st.plotly_chart(fig_scenarios, use_container_width=True)
            
            optimal_scenario = extended_analysis_results['optimal_scenario']
            if optimal_scenario["variation"] != 0:
                st.info(f"""
                💡 **Suggerimento tariffario ottimale**: 
                Un'ADR di €{optimal_scenario['adr_lordo']:.2f} ({optimal_scenario['variation']:+}%) 
                genererebbe un profitto totale di €{optimal_scenario['total_rev_profit']:,.2f}, 
                con un incremento di €{optimal_scenario['total_rev_profit'] - metrics['total_rev_profit']:,.2f} 
                rispetto all'ADR proposta originalmente.
                """)
            else:
                st.success("✅ L'ADR proposta è già ottimale per massimizzare il profitto.")
            
            if data_source == "Import file Excel" and 'raw_excel_data' in st.session_state and 'critical_days' in extended_analysis_results:
                st.subheader("Analisi Shoulder Days")
                
                st.dataframe(
                    extended_analysis_results['critical_days'],
                    column_config={
                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                        "giorno": "Giorno",
                        "finale_rn": st.column_config.NumberColumn("OTB Forecast", format="%d"),
                        "camere_gruppo": st.column_config.NumberColumn("Richiesta", format="%d"),
                        "camere_displaced": st.column_config.NumberColumn("Displaced", format="%d"),
                        "criticità": st.column_config.SelectboxColumn(
                            "Criticità",
                            options=["Nessuna", "Bassa", "Media", "Alta"],
                            required=True
                        )
                    }
                )
                
                st.markdown("*La funzionalità di suggerimento date alternative sarà disponibile nei prossimi aggiornamenti.*")
               
        st.subheader("Dati Dettagliati")
        display_cols = ['data', 'giorno', 'finale_rn', 'camere_gruppo', 'camere_disponibili', 
                         'camere_displaced', 'adr_gruppo_netto', 'finale_adr', 
                         'revenue_camere_gruppo_effettivo', 'revenue_displaced', 'impatto_revenue_totale']
           
        st.dataframe(
               result_df[display_cols],
               column_config={
                   "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                   "giorno": "Giorno",
                   "finale_rn": st.column_config.NumberColumn("FCST OTB", format="%d"),
                   "camere_gruppo": st.column_config.NumberColumn("REQ", format="%d"),
                   "camere_disponibili": st.column_config.NumberColumn("Disponibili", format="%d"),
                   "camere_displaced": st.column_config.NumberColumn("DSPL", format="%d"),
                   "adr_gruppo_netto": st.column_config.NumberColumn("ADR Netto", format="€%.2f"),
                   "finale_adr": st.column_config.NumberColumn("ADR Attuale", format="€%.2f"),
                   "revenue_camere_gruppo_effettivo": st.column_config.NumberColumn("REV REQ", format="€%.2f"),
                   "revenue_displaced": st.column_config.NumberColumn("REV DSPL", format="€%.2f"),
                   "impatto_revenue_totale": st.column_config.NumberColumn("DIFF", format="€%.2f")
               }
           )
           
        st.markdown(get_csv_download_link(result_df, f"displacement_{group_name}", "📥 Scarica dati completi (CSV)"), unsafe_allow_html=True)
           
        st.header("Decisione Finale")
           
        decision_color = COLOR_PALETTE["positive"] if metrics['should_accept'] else COLOR_PALETTE["negative"]
        decision_text = "ACCETTA GRUPPO" if metrics['should_accept'] else "DECLINA GRUPPO"
           
        st.markdown(f"""
        <div style="background-color:{decision_color}; padding:20px; border-radius:10px; text-align:center; margin-top:20px;">
               <h2 style="color:white; margin:0;">{decision_text}</h2>
               <p style="color:white; margin-top:10px;">
                   Impatto Revenue: €{metrics['total_impact']:,.2f} | 
                   ADR Netto: €{metrics['current_adr_netto']:.2f} | 
                   Camere: {metrics['accepted_rooms']}/{metrics['total_group_rooms']} |
                   Displacement: {metrics['displaced_rooms']} camere
               </p>
           </div>
        """, unsafe_allow_html=True)
           
        if metrics['needs_authorization'] and not enable_series:
            st.warning("⚠️ Questa richiesta gruppo supera il valore di €35.000 e richiede autorizzazione")
               
            st.subheader("Email di Richiesta Autorizzazione")
               
            email_text = generate_auth_email(
                group_name=group_name,
                total_revenue=metrics['total_lordo'],
                dates=result_df['data'].tolist(),
                rooms=num_rooms,
                adr=adr_lordo
            )
               
            st.text_area("Email da inviare", email_text, height=300, key="email_text")
               
            col1, col2 = st.columns([1, 2])
            with col1:
                st.download_button(
                    label="📥 Scarica Email",
                    data=email_text,
                    file_name=f"richiesta_autorizzazione_{group_name}.txt",
                    mime="text/plain"
                )
            with col2:
                st.info("💡 Per copiare l'email, seleziona tutto il testo nella casella sopra (Ctrl+A), poi premi Ctrl+C (o Cmd+C su Mac)")
        
        if enable_series:
            if st.button("Salva passaggio e continua", key="save_passage"):
                st.session_state['series_data'].append({
                    'passage': st.session_state['current_passage'],
                    'date_range': f"{group_arrival.strftime('%d/%m/%Y')} - {group_departure.strftime('%d/%m/%Y')}",
                    'rooms': num_rooms,
                    'adr': adr_lordo,
                    'room_revenue': metrics['group_room_revenue'],
                    'ancillary_revenue': metrics['group_ancillary'],
                    'total_revenue': metrics['group_room_revenue'] + metrics['group_ancillary'],
                    'total_lordo': metrics['total_lordo'],
                    'displaced_revenue': metrics['revenue_displaced'],
                    'net_impact': metrics['total_impact'],
                    'analysis_data': result_df.copy()
                })
                
                if st.session_state['current_passage'] < num_passages:
                    st.session_state['current_passage'] += 1
                    st.session_state['analysis_phase'] = 'start'
                    
                    next_start_date = group_departure
                    next_end_date = next_start_date + timedelta(days=date_nights)
                    
                    st.session_state['next_start_date'] = next_start_date
                    st.session_state['next_end_date'] = next_end_date
                else:
                    st.session_state['series_complete'] = True
                st.rerun()
        else:
            if st.button("Nuova Analisi", key="new_analysis"):
                st.session_state['analysis_phase'] = 'start'
                st.rerun()

if enable_series and st.session_state.get('series_complete', False):
    st.header("Riepilogo Serie di Gruppi")
    
    series_df_data = []
    for item in st.session_state['series_data']:
        series_df_data.append({
            'Passaggio': item['passage'],
            'Date': item['date_range'],
            'Camere': item['rooms'],
            'ADR': item['adr'],
            'Rev. Camere': item['room_revenue'],
            'Rev. Ancillare': item['ancillary_revenue'],
            'DSPL': item['displaced_revenue'],
            'Impatto': item['net_impact'],
            'Totale Lordo': item['total_lordo']
        })
    
    series_df = pd.DataFrame(series_df_data)
    
    st.dataframe(
        series_df,
        column_config={
            "Passaggio": st.column_config.NumberColumn("Passaggio", format="%d"),
            "Date": "Date",
            "Camere": st.column_config.NumberColumn("Camere", format="%d"),
            "ADR": st.column_config.NumberColumn("ADR", format="€%.2f"),
            "Rev. Camere": st.column_config.NumberColumn("Rev. Camere", format="€%.2f"),
            "Rev. Ancillare": st.column_config.NumberColumn("Rev. Ancillare", format="€%.2f"),
            "DSPL": st.column_config.NumberColumn("DSPL", format="€%.2f"),
            "Impatto": st.column_config.NumberColumn("Impatto", format="€%.2f"),
            "Totale Lordo": st.column_config.NumberColumn("Totale Lordo", format="€%.2f")
        }
    )
    
    total_series_revenue = sum(item['total_lordo'] for item in st.session_state['series_data'])
    total_series_impact = sum(item['net_impact'] for item in st.session_state['series_data'])
    total_series_room_revenue = sum(item['room_revenue'] for item in st.session_state['series_data'])
    total_series_ancillary = sum(item['ancillary_revenue'] for item in st.session_state['series_data'])
    total_series_displaced = sum(item['displaced_revenue'] for item in st.session_state['series_data'])
    
    st.subheader("Totale Serie")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Revenue Totale Serie (lordo)", f"€{total_series_revenue:,.2f}")
    with col2:
        st.metric("Revenue Perso", f"€{total_series_displaced:,.2f}")
    with col3:
        st.metric("Impatto Totale Serie", f"€{total_series_impact:,.2f}")
    
    fig = px.bar(
        series_df, 
        x='Passaggio', 
        y=['Rev. Camere', 'Rev. Ancillare', 'DSPL'], 
        barmode='group', 
        title="Confronto Revenue per Passaggio",
        labels={'value': 'Revenue (€)', 'variable': 'Categoria'},
        color_discrete_map={
            'Rev. Camere': COLOR_PALETTE["secondary"],
            'Rev. Ancillare': COLOR_PALETTE["primary"],
            'DSPL': COLOR_PALETTE["negative"]
        }
    )
    
    fig.update_layout(
        font_family="Inter, sans-serif",
        plot_bgcolor=COLOR_PALETTE["background"],
        paper_bgcolor=COLOR_PALETTE["background"],
        font_color=COLOR_PALETTE["text"]
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    if total_series_revenue > 35000:
        st.warning("⚠️ Questa serie di gruppi supera il valore di €35.000 e richiede autorizzazione")
        
        st.subheader("Email di Richiesta Autorizzazione (Serie)")
        
        series_summary_for_email = pd.DataFrame({
            'Passaggio': series_df['Passaggio'],
            'Date': series_df['Date'],
            'Camere': series_df['Camere'],
            'ADR': series_df['ADR'].map('€{:.2f}'.format),
            'Totale': series_df['Totale Lordo'].map('€{:.2f}'.format),
            'Impatto': series_df['Impatto'].map('€{:.2f}'.format)
        })
        
        email_text = generate_series_auth_email(
            group_name=group_name,
            total_series_revenue=total_series_revenue,
            total_series_impact=total_series_impact,
            num_passages=num_passages,
            series_summary=series_summary_for_email
        )
        
        st.text_area("Email da inviare", email_text, height=400, key="email_series")
        
        col1, col2 = st.columns([1, 2])
        with col1:
            st.download_button(
                label="📥 Scarica Email",
                data=email_text,
                file_name=f"richiesta_autorizzazione_serie_{group_name}.txt",
                mime="text/plain"
            )
        with col2:
            st.info("💡 Per copiare l'email, seleziona tutto il testo nella casella sopra (Ctrl+A), poi premi Ctrl+C (o Cmd+C su Mac)")
    
    if st.button("Nuova Serie", key="new_series"):
        for key in ['series_data', 'current_passage', 'series_complete']:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state['analysis_phase'] = 'start'
        st.rerun()

st.markdown("---")
st.markdown(
   f"""
   <div style='text-align: center; font-family: Inter, sans-serif; color: #5E5E5E; font-size: 0.8rem;'>
       <p>Hotel Group Displacement Analyzer | v0.9.0beta2 developed by Alessandro Merella | Original excel concept and formulas by Andrea Conte<br>
       Sessione: {st.session_state['username']} | Ultimo accesso: {datetime.fromtimestamp(st.session_state['login_time']).strftime('%d/%m/%Y %H:%M')}<br>
       Distributed under MIT License
       </p>
   </div>
   """, 
   unsafe_allow_html=True
)
