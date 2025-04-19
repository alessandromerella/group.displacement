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
import xlsxwriter
import traceback

st.set_page_config(page_title="Hotel Groups Displacement Analyzer v0.9.5r5", layout="wide")

COLOR_PALETTE = {
    "primary": "#D8C0B7",
    "secondary": "#8CA68C",
    "text": "#5E5E5E",
    "background": "#F8F6F4",
    "accent": "#B6805C",
    "positive": "#8CA68C",
    "negative": "#D8837F"
}

js_code = """
<script>
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(function() {
        const editableCells = document.querySelectorAll('[data-baseweb="data-grid"] [role="gridcell"]:not([aria-readonly="true"])');
        
        editableCells.forEach((cell, index) => {
            cell.setAttribute('tabindex', index + 1);
        });
        
        document.addEventListener('keydown', function(event) {
            if (event.key === 'Tab') {
                const activeElement = document.activeElement;
                
                if (activeElement && activeElement.getAttribute('role') === 'gridcell') {
                    const currentIndex = parseInt(activeElement.getAttribute('tabindex'));
                    let nextIndex;
                    
                    if (event.shiftKey) {
                        nextIndex = currentIndex - 1;
                        if (nextIndex < 1) {
                            nextIndex = editableCells.length;
                        }
                    } else {
                        nextIndex = currentIndex + 1;
                        if (nextIndex > editableCells.length) {
                            nextIndex = 1;
                        }
                    }
                    
                    const nextCell = document.querySelector(`[data-baseweb="data-grid"] [role="gridcell"][tabindex="${nextIndex}"]`);
                    if (nextCell) {
                        event.preventDefault();
                        nextCell.focus();
                        
                        const clickEvent = new MouseEvent('dblclick', {
                            bubbles: true,
                            cancelable: true,
                            view: window
                        });
                        nextCell.dispatchEvent(clickEvent);
                    }
                }
            }
        });
    }, 1000);
});
</script>
"""

def parse_booking_request(text):
    results = {
        'group_name': None,
        'arrival_date': None,
        'departure_date': None,
        'num_rooms': None,
        'is_checkout': True
    }
    
    agency_patterns = [
        r'nome\s+agenzia\s*:\s*([^\n]+)',
        r'agenzia\s*:\s*([^\n]+)',
        r'gruppo\s*:\s*([^\n]+)',
        r'nome\s+gruppo\s*:\s*([^\n]+)',
        r'agenzia[\s:]+([^\n,]+)'
    ]
    
    for pattern in agency_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            results['group_name'] = match.group(1).strip()
            break
    
    periodo_pattern = r'PERIODO:\s+dal\s+(\d+)\s+([a-zA-Z]+)\s+al\s+(\d+)\s+([a-zA-Z]+)\s+(\d{4})'
    periodo_match = re.search(periodo_pattern, text, re.IGNORECASE)
    
    if periodo_match:
        day1, month1, day2, month2, year = periodo_match.groups()
        
        month_map = {
            'gennaio': 1, 'febbraio': 2, 'marzo': 3, 'aprile': 4, 'maggio': 5, 'giugno': 6,
            'luglio': 7, 'agosto': 8, 'settembre': 9, 'ottobre': 10, 'novembre': 11, 'dicembre': 12,
            'gen': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'mag': 5, 'giu': 6,
            'lug': 7, 'ago': 8, 'set': 9, 'ott': 10, 'nov': 11, 'dic': 12
        }
        
        month_num1 = month_map.get(month1.lower(), None)
        month_num2 = month_map.get(month2.lower(), None)
        
        if month_num1 and month_num2:
            try:
                results['arrival_date'] = datetime(int(year), month_num1, int(day1)).date()
                results['departure_date'] = datetime(int(year), month_num2, int(day2)).date()
            except ValueError:
                pass
    
    if results['arrival_date'] is None:
        date_pattern = r'[Dd]al\s+(\d+)\s+([a-zA-Z]+)\s+(?:al|a)\s+(\d+)\s+([a-zA-Z]+)\s+(\d{4})'
        match = re.search(date_pattern, text)
        if match:
            day1, month1, day2, month2, year = match.groups()
            
            month_map = {
                'gennaio': 1, 'febbraio': 2, 'marzo': 3, 'aprile': 4, 'maggio': 5, 'giugno': 6,
                'luglio': 7, 'agosto': 8, 'settembre': 9, 'ottobre': 10, 'novembre': 11, 'dicembre': 12,
                'gen': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'mag': 5, 'giu': 6,
                'lug': 7, 'ago': 8, 'set': 9, 'ott': 10, 'nov': 11, 'dic': 12
            }
            
            month_num1 = month_map.get(month1.lower(), None)
            month_num2 = month_map.get(month2.lower(), None)
            
            if month_num1 and month_num2:
                try:
                    results['arrival_date'] = datetime(int(year), month_num1, int(day1)).date()
                    
                    if "incluso" in text.lower():
                        results['is_checkout'] = False
                        results['departure_date'] = datetime(int(year), month_num2, int(day2) + 1).date()
                    else:
                        results['departure_date'] = datetime(int(year), month_num2, int(day2)).date()
                except ValueError:
                    pass
    
    room_patterns = [
        r'(\d+)\s+camer[ae]',
        r'camer[ae]\s*:\s*(\d+)',
        r'n\.\s*camer[ae]\s*:\s*(\d+)'
    ]
    
    for pattern in room_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            results['num_rooms'] = int(match.group(1))
            break
    
    return results

def load_changelog():
    try:
        with open("changelog.md", "r") as f:
            return f.read()
    except FileNotFoundError:
        return """# Changelog Hotel Group Displacement Analyzer

## v0.9.5r5 (Attuale)
- **Bugfix**: Risolto problema DatetimeIndex in analisi tipologie camere
- **Bugfix**: Migliorata gestione delle date nel parser richieste booking 
- **Miglioramento UX**: Correzione persistenza dati tra sessioni
- **Performance**: Ottimizzazione gestione errori e robustezza generale

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
    
    st.markdown("""
    <div style="display: flex; justify-content: center; align-items: center; flex-direction: column; margin-bottom: 20px;">
        <img src="https://www.revguardian.altervista.org/hgd.logo.png" style="width: 200px; margin-bottom: 10px;">
        <p style="text-align: center; margin: 0;">v0.9.5r5</p>
        <p style="text-align: center; margin-top: 10px;">Accedi per continuare</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 1, 2])
    with col2:
        if st.button("📋 What's New", key="whats_new_btn"):
            st.session_state['show_changelog'] = True
    
    if 'show_changelog' in st.session_state and st.session_state['show_changelog']:
        try:
            with st.expander("Changelog - Novità della versione", expanded=True):
                changelog_content = load_changelog()
                st.markdown(changelog_content)
                
                if st.button("Chiudi", key="close_changelog"):
                    st.session_state['show_changelog'] = False
                    st.rerun()
        except Exception as e:
            st.error(f"Errore durante la visualizzazione del changelog: {e}")
    
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

st.sidebar.info(f"Accesso effettuato come: {st.session_state['username']}")
if st.sidebar.button("Logout"):
    for key in ['authenticated', 'username', 'login_time', 'analysis_phase', 'wizard_step', 'forecast_method', 
               'pickup_factor', 'pickup_percentage', 'pickup_value', 'series_data', 'current_passage', 
               'series_complete', 'raw_excel_data', 'available_dates', 'analyzed_data', 'selected_start_date', 
               'selected_end_date', 'events_data_cache', 'events_data_updated', 'enable_extended_reasoning',
               'room_types_df', 'rooms_by_day_df', 'booking_data', 'force_update', 'default_start_date', 
               'default_end_date', 'booking_data_json', 'update_pending', 'group_name_stored', 'num_rooms_stored', 
               'arrival_date_stored', 'departure_date_stored']:
        if key in st.session_state:
            del st.session_state[key]
    changelog_keys = [k for k in st.session_state if k.startswith("has_seen_changelog_")]
    for key in changelog_keys:
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
    
    [data-baseweb="data-grid"] [role="gridcell"]:not([aria-readonly="true"]) {{
        background-color: rgba(140, 166, 140, 0.1);
        font-weight: bold;
        font-size: 1.05em;
    }}
    
    [data-baseweb="data-grid"] [role="gridcell"]:focus {{
        background-color: rgba(140, 166, 140, 0.2);
        outline: 2px solid {COLOR_PALETTE["accent"]};
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
    
    .download-button {{
        display: inline-block;
        padding: 0.5em 1em;
        background-color: #4CAF50;
        color: white;
        text-align: center;
        text-decoration: none;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }}
    .download-button:hover {{
        background-color: #45a049;
    }}
</style>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)

st.markdown(js_code, unsafe_allow_html=True)

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

def safe_date_conversion(date_val):
    if pd.isna(date_val):
        return pd.NaT
    
    if isinstance(date_val, (datetime, date)):
        return pd.to_datetime(date_val)
    
    date_str = str(date_val)
    date_only = re.search(r'(\d{2}/\d{2}/\d{4})', date_str)
    if date_only:
        try:
            return pd.to_datetime(date_only.group(1), format='%d/%m/%Y')
        except:
            pass
    
    formats = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y']
    for fmt in formats:
        try:
            return pd.to_datetime(date_str, format=fmt, errors='coerce')
        except:
            continue
    
    return pd.NaT

def process_excel_import(uploaded_files):
    st.write(f"Debug: Numero file caricati: {len(uploaded_files)}")
    for file in uploaded_files:
        st.write(f"Debug: Nome file: {file.name}")
    
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
            st.write(f"Debug: Prime righe di {uploaded_file.name}:")
            st.write(df.head())
            
            file_type, year, month_year = identify_excel_file_type(df)
            st.write(f"Debug: Tipo file: {file_type}, Anno: {year}")
            
            date_col_candidates = ['Giorno', 'Data', 'Date', 'day', 'date']
            data_rows = None
            
            for candidate in date_col_candidates:
                mask = df.iloc[:, 0].astype(str).str.contains(candidate, case=False, na=False)
                if mask.any():
                    data_rows = df[mask].index[0]
                    break
            
            if data_rows is None:
                st.error(f"Formato del file {uploaded_file.name} non riconosciuto: nessuna colonna data trovata")
                continue
                
            headers = df.iloc[data_rows].values.tolist()
            
            data_df = df.iloc[data_rows+1:].reset_index(drop=True)
            data_df.columns = headers[:len(data_df.columns)]
            
            date_column = None
            for col in data_df.columns:
                if isinstance(col, str) and any(candidate.lower() in col.lower() for candidate in date_col_candidates):
                    date_column = col
                    break
            
            if date_column is None:
                date_column = data_df.columns[0]
            
            data_df = data_df.rename(columns={date_column: 'Giorno'})
            
            data_df = data_df[~data_df['Giorno'].isna()]
            data_df = data_df[~data_df['Giorno'].astype(str).str.contains('Filtri applicati:', na=False)]
            
            data_df['Giorno'] = data_df['Giorno'].apply(safe_date_conversion)
            data_df = data_df.dropna(subset=['Giorno'])
            
            if len(data_df) == 0:
                st.error(f"Il file {uploaded_file.name} non contiene date valide dopo la pulizia")
                continue
            
            st.write(f"Debug: Date estratte da {uploaded_file.name}:")
            st.write(data_df['Giorno'].head())
            
            numeric_cols = ['Room nights', 'Bed nights', 'ADR Cam', 'ADR Bed', 'Room Revenue', 'RevPar']
            for col in numeric_cols:
                if col in data_df.columns:
                    data_df[col] = pd.to_numeric(data_df[col], errors='coerce').fillna(0)
            
            data_df.rename(columns={'Giorno': 'data'}, inplace=True)
            
            if file_type == "IDV":
                if str(current_year) in str(year) or (month_year and str(current_year) in str(month_year)):
                    idv_cy_data = data_df
                    st.success(f"File IDV Anno Corrente riconosciuto: {uploaded_file.name}")
                else:
                    idv_ly_data = data_df
                    st.success(f"File IDV Anno Precedente riconosciuto: {uploaded_file.name}")
            elif file_type == "GRP":
                if 'Room nights' in data_df.columns and data_df['Room nights'].sum() > 0:
                    grp_otb_data = data_df
                    st.success(f"File Gruppi Confermati riconosciuto: {uploaded_file.name}")
                else:
                    grp_opz_data = data_df
                    st.success(f"File Gruppi Opzionati riconosciuto: {uploaded_file.name}")
        
        except Exception as e:
            st.error(f"Errore nell'elaborazione del file {uploaded_file.name}: {str(e)}")
            st.code(traceback.format_exc())
    
    try:
        has_valid_dates = (idv_cy_data is not None and 'data' in idv_cy_data.columns 
                          and len(idv_cy_data) > 0 and not idv_cy_data['data'].isna().all())
        
        if has_valid_dates:
            min_date = idv_cy_data['data'].min()
            max_date = idv_cy_data['data'].max()
            
            st.write(f"Debug: Range date trovate: {min_date} - {max_date}")
            
            if pd.notna(min_date) and pd.notna(max_date):
                available_dates = pd.date_range(start=min_date, end=max_date)
                st.session_state['available_dates'] = available_dates
                st.success(f"File elaborati con successo! Range date disponibili: {min_date.strftime('%d/%m/%Y')} - {max_date.strftime('%d/%m/%Y')}")
            else:
                raise ValueError("Date non valide nei dati importati")
        else:
            raise ValueError("Dati non validi o colonna data mancante")
    except Exception as e:
        st.error(f"Impossibile determinare il range di date: {str(e)}")
        fallback_start = datetime.now()
        fallback_end = fallback_start + timedelta(days=30)
        available_dates = pd.date_range(start=fallback_start, end=fallback_end)
        st.session_state['available_dates'] = available_dates
        st.warning(f"Utilizzando date predefinite: {fallback_start.strftime('%d/%m/%Y')} - {fallback_end.strftime('%d/%m/%Y')}")
    
    if idv_cy_data is not None and len(idv_cy_data) > 0:
        st.write("Debug: Dati IDV CY trovati")
        st.write(idv_cy_data.head())
    if idv_ly_data is not None and len(idv_ly_data) > 0:
        st.write("Debug: Dati IDV LY trovati")
        st.write(idv_ly_data.head())
    
    return idv_cy_data, idv_ly_data, grp_otb_data, grp_opz_data

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

def process_imported_data(idv_cy_data, idv_ly_data, grp_otb_data, grp_opz_data, date_range):
    try:
        st.write(f"Debug: Elaborazione date range: {date_range.min()} - {date_range.max()}")
        result_dates = date_range
        result_df = pd.DataFrame({'data': result_dates})
        
        result_df['giorno'] = result_df['data'].dt.strftime('%a')
        result_df['data_ly'] = result_df['data'].apply(same_day_last_year)
        result_df['giorno_ly'] = result_df['data_ly'].dt.strftime('%a')
        
        if idv_cy_data is not None and 'data' in idv_cy_data.columns:
            if 'Room nights' in idv_cy_data.columns:
                idv_cy_data.rename(columns={
                    'Room nights': 'otb_ind_rn',
                    'ADR Cam': 'otb_ind_adr'
                }, inplace=True)
            
            st.write("Debug: Filtro date IDV CY")
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
        else:
            result_df['otb_ind_rn'] = 0
            result_df['otb_ind_adr'] = 0
        
        if idv_ly_data is not None and 'data' in idv_ly_data.columns:
            if 'Room nights' in idv_ly_data.columns:
                idv_ly_data.rename(columns={
                    'Room nights': 'ly_ind_rn',
                    'ADR Cam': 'ly_ind_adr'
                }, inplace=True)
            
            idv_ly_data['data'] = pd.to_datetime(idv_ly_data['data'])
            
            st.write("Debug: Filtro date IDV LY")
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
        else:
            result_df['ly_ind_rn'] = 0
            result_df['ly_ind_adr'] = 0
        
        if grp_otb_data is not None and not grp_otb_data.empty and 'data' in grp_otb_data.columns:
            if 'Room nights' in grp_otb_data.columns:
                grp_otb_data.rename(columns={
                    'Room nights': 'grp_otb_rn',
                    'ADR Cam': 'grp_otb_adr'
                }, inplace=True)
            
            st.write("Debug: Filtro date GRP OTB")
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
        
        if grp_opz_data is not None and not grp_opz_data.empty and 'data' in grp_opz_data.columns:
            if 'Room nights' in grp_opz_data.columns:
                grp_opz_data.rename(columns={
                    'Room nights': 'grp_opz_rn',
                    'ADR Cam': 'grp_opz_adr'
                }, inplace=True)
            
            st.write("Debug: Filtro date GRP OPZ")
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
        
        if 'forecast_method' not in st.session_state:
            st.session_state['forecast_method'] = "LY - OTB"
        
        if st.session_state['forecast_method'] == "Basato su LY":
            result_df['fcst_ind_rn'] = np.ceil(result_df['ly_ind_rn'] * st.session_state.get('pickup_factor', 1.0))
        elif st.session_state['forecast_method'] == "Percentuale su OTB":
            result_df['fcst_ind_rn'] = np.ceil(result_df['otb_ind_rn'] * (1 + st.session_state.get('pickup_percentage', 20)/100))
        elif st.session_state['forecast_method'] == "Valore assoluto":
            result_df['fcst_ind_rn'] = result_df['otb_ind_rn'] + st.session_state.get('pickup_value', 10)
        else:
            result_df['fcst_ind_rn'] = result_df['ly_ind_rn'] - result_df['otb_ind_rn']
            result_df['fcst_ind_rn'] = result_df['fcst_ind_rn'].apply(lambda x: max(0, x))
        
        result_df['fcst_ind_adr'] = result_df['otb_ind_adr']
        
        result_df['otb_ind_rev'] = result_df['otb_ind_rn'] * result_df['otb_ind_adr']
        result_df['ly_ind_rev'] = result_df['ly_ind_rn'] * result_df['ly_ind_adr']
        result_df['grp_otb_rev'] = result_df['grp_otb_rn'] * result_df['grp_otb_adr']
        result_df['grp_opz_rev'] = result_df['grp_opz_rn'] * result_df['grp_opz_adr']
        result_df['fcst_ind_rev'] = result_df['fcst_ind_rn'] * result_df['fcst_ind_adr']
        
        result_df['finale_rn'] = result_df['fcst_ind_rn'] + result_df['otb_ind_rn'] + result_df['grp_otb_rn']
        result_df['finale_opz_rn'] = result_df['finale_rn'] + result_df['grp_opz_rn']
        
        result_df['finale_rev'] = result_df['otb_ind_rev'] + result_df['fcst_ind_rev'] + result_df['grp_otb_rev']
        result_df['finale_adr'] = np.where(result_df['finale_rn'] > 0,
                                       result_df['finale_rev'] / result_df['finale_rn'],
                                       0)
        
        st.write("Debug: Risultato finale elaborazione")
        st.write(result_df.head())
        
        return result_df
        
    except Exception as e:
        st.error(f"Errore nell'elaborazione dei dati importati: {e}")
        st.code(traceback.format_exc())
        return None

def get_booking_data():
    if 'booking_data_json' in st.session_state:
        try:
            data = json.loads(st.session_state['booking_data_json'])
            
            if data.get('arrival_date'):
                data['arrival_date'] = datetime.strptime(data['arrival_date'], '%Y-%m-%d').date()
            if data.get('departure_date'):
                data['departure_date'] = datetime.strptime(data['departure_date'], '%Y-%m-%d').date()
                
            return data
        except Exception as e:
            st.error(f"Errore nel recupero dei dati booking: {e}")
            return None
    return None

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

def generate_excel_report(result_df, metrics, group_info, hotel_info):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': COLOR_PALETTE["background"],
        'font_color': COLOR_PALETTE["text"],
        'border': 1
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': COLOR_PALETTE["primary"],
        'font_color': 'white',
        'border': 1
    })
    
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    number_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '#,##0'
    })
    
    currency_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '€#,##0.00'
    })
    
    percentage_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0.0%'
    })
    
    date_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': 'dd/mm/yyyy'
    })
    
    section_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': COLOR_PALETTE["secondary"],
        'font_color': 'white',
        'border': 1
    })
    
    result_positive = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': COLOR_PALETTE["positive"],
        'font_color': 'white',
        'border': 1
    })
    
    result_negative = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': COLOR_PALETTE["negative"],
        'font_color': 'white',
        'border': 1
    })
    
    summary_sheet = workbook.add_worksheet('Riepilogo')
    summary_sheet.set_column('A:A', 30)
    summary_sheet.set_column('B:B', 30)
    summary_sheet.set_column('C:C', 20)
    summary_sheet.set_column('D:D', 20)
    
    summary_sheet.merge_range('A1:D1', 'RIEPILOGO ANALISI DISPLACEMENT', title_format)
    
    summary_sheet.merge_range('A3:D3', 'DATI HOTEL', section_format)
    summary_sheet.write('A4', 'Hotel', cell_format)
    summary_sheet.write('B4', hotel_info['name'], cell_format)
    summary_sheet.write('A5', 'Capacità camere', cell_format)
    summary_sheet.write('B5', hotel_info['capacity'], number_format)
    summary_sheet.write('A6', 'Aliquota IVA', cell_format)
    summary_sheet.write('B6', hotel_info['iva_rate'], percentage_format)
    
    summary_sheet.merge_range('A8:D8', 'DATI GRUPPO', section_format)
    summary_sheet.write('A9', 'Nome gruppo', cell_format)
    summary_sheet.write('B9', group_info['name'], cell_format)
    summary_sheet.write('A10', 'Data arrivo', cell_format)
    summary_sheet.write('B10', group_info['arrival_date'], date_format)
    summary_sheet.write('A11', 'Data partenza', cell_format)
    summary_sheet.write('B11', group_info['departure_date'], date_format)
    summary_sheet.write('A12', 'Numero camere', cell_format)
    summary_sheet.write('B12', group_info['num_rooms'], number_format)
    summary_sheet.write('A13', 'ADR lordo', cell_format)
    summary_sheet.write('B13', group_info['adr_lordo'], currency_format)
    summary_sheet.write('A14', 'ADR netto', cell_format)
    summary_sheet.write('B14', group_info['adr_netto'], currency_format)
    summary_sheet.write('A15', 'Revenue ancillare', cell_format)
    summary_sheet.write('B15', group_info['ancillary_revenue'], currency_format)
    
    summary_sheet.merge_range('A17:D17', 'RISULTATI ANALISI', section_format)
    summary_sheet.write('A18', 'Camere richieste', cell_format)
    summary_sheet.write('B18', metrics['total_group_rooms'], number_format)
    summary_sheet.write('A19', 'Camere accettate', cell_format)
    summary_sheet.write('B19', metrics['accepted_rooms'], number_format)
    summary_sheet.write('A20', 'Camere displaced', cell_format)
    summary_sheet.write('B20', metrics['displaced_rooms'], number_format)
    summary_sheet.write('A21', 'Revenue camere gruppo', cell_format)
    summary_sheet.write('B21', metrics['group_room_revenue'], currency_format)
    summary_sheet.write('A22', 'Revenue ancillare', cell_format)
    summary_sheet.write('B22', metrics['group_ancillary'], currency_format)
    summary_sheet.write('A23', 'Revenue displaced', cell_format)
    summary_sheet.write('B23', metrics['revenue_displaced'], currency_format)
    summary_sheet.write('A24', 'Impatto totale', cell_format)
    summary_sheet.write('B24', metrics['total_impact'], currency_format)
    summary_sheet.write('A25', 'Valore totale lordo', cell_format)
    summary_sheet.write('B25', metrics['total_lordo'], currency_format)
    
    decision_text = "ACCETTA GRUPPO" if metrics['should_accept'] else "DECLINA GRUPPO"
    decision_format = result_positive if metrics['should_accept'] else result_negative
    summary_sheet.merge_range('A27:D27', decision_text, decision_format)
    
    if metrics['needs_authorization']:
        summary_sheet.merge_range('A29:D29', 'ATTENZIONE: RICHIEDE AUTORIZZAZIONE (>€35.000)', result_negative)
    
    summary_sheet.merge_range('A31:D31', f'Report generato il {datetime.now().strftime("%d/%m/%Y %H:%M")} da {st.session_state["username"]}', cell_format)
    
    data_sheet = workbook.add_worksheet('Dati Dettagliati')
    
    columns = ['data', 'giorno', 'finale_rn', 'camere_gruppo', 'camere_disponibili', 
              'camere_displaced', 'adr_gruppo_netto', 'finale_adr', 
              'revenue_camere_gruppo_effettivo', 'revenue_displaced', 'impatto_revenue_totale']
    
    headers = ['Data', 'Giorno', 'FCST OTB', 'REQ', 'Disponibili', 
              'DSPL', 'ADR Netto', 'ADR Attuale', 
              'REV REQ', 'REV DSPL', 'DIFF']
    
    for i, col in enumerate(columns):
        data_sheet.set_column(i, i, 15)
    
    for i, header in enumerate(headers):
        data_sheet.write(0, i, header, header_format)
    
    for i, row in result_df.reset_index(drop=True).iterrows():
        data_sheet.write(i+1, 0, row['data'], date_format)
        data_sheet.write(i+1, 1, row['giorno'], cell_format)
        data_sheet.write(i+1, 2, row['finale_rn'], number_format)
        data_sheet.write(i+1, 3, row['camere_gruppo'], number_format)
        data_sheet.write(i+1, 4, row['camere_disponibili'], number_format)
        data_sheet.write(i+1, 5, row['camere_displaced'], number_format)
        data_sheet.write(i+1, 6, row['adr_gruppo_netto'], currency_format)
        data_sheet.write(i+1, 7, row['finale_adr'], currency_format)
        data_sheet.write(i+1, 8, row['revenue_camere_gruppo_effettivo'], currency_format)
        data_sheet.write(i+1, 9, row['revenue_displaced'], currency_format)
        data_sheet.write(i+1, 10, row['impatto_revenue_totale'], currency_format)
    
    forecast_sheet = workbook.add_worksheet('Forecast e OTB')
    
    forecast_cols = ['data', 'giorno', 'otb_ind_rn', 'ly_ind_rn', 'fcst_ind_rn', 
                    'grp_otb_rn', 'grp_opz_rn', 'finale_rn']
    
    forecast_headers = ['Data', 'Giorno', 'OTB IND', 'LY IND', 'FCST IND', 
                        'GRP OTB', 'GRP OPZ', 'TOTALE']
    
    for i, col in enumerate(forecast_cols):
        forecast_sheet.set_column(i, i, 15)
    
    for i, header in enumerate(forecast_headers):
        forecast_sheet.write(0, i, header, header_format)
    
    for i, row in result_df.reset_index(drop=True).iterrows():
        forecast_sheet.write(i+1, 0, row['data'], date_format)
        forecast_sheet.write(i+1, 1, row['giorno'], cell_format)
        forecast_sheet.write(i+1, 2, row['otb_ind_rn'], number_format)
        forecast_sheet.write(i+1, 3, row['ly_ind_rn'], number_format)
        forecast_sheet.write(i+1, 4, row['fcst_ind_rn'], number_format)
        forecast_sheet.write(i+1, 5, row['grp_otb_rn'], number_format)
        forecast_sheet.write(i+1, 6, row['grp_opz_rn'], number_format)
        forecast_sheet.write(i+1, 7, row['finale_rn'], number_format)
    
    workbook.close()
    
    output.seek(0)
    return output.getvalue()

def get_excel_download_link(result_df, metrics, group_info, hotel_info, filename):
    excel_data = generate_excel_report(result_df, metrics, group_info, hotel_info)
    b64 = base64.b64encode(excel_data).decode()
    
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx" class="download-button">📥 Scarica Report Excel</a>'

def generate_auth_email(group_name, total_revenue, dates, rooms, adr, nights):
    email_template = f"""
    Oggetto: Richiesta autorizzazione gruppo {group_name} - Valore €{total_revenue:,.2f}
    
    Gentile Revenue Manager,
    
    Richiedo autorizzazione per l'offerta al gruppo "{group_name}" che supera la soglia di €35.000.
    
    Dettagli della richiesta:
    - Date soggiorno: dal {dates[0].strftime('%d/%m/%Y')} al {dates[-1].strftime('%d/%m/%Y')}
    - Numero camere: {rooms} ROH
    - Numero notti: {nights}
    - ADR: €{adr:.2f}
    - Valore totale: €{total_revenue:,.2f} ({rooms} camere × €{adr:.2f} × {nights} notti + eventuali ancillari)
    
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

class ExcelCompatibleDisplacementAnalyzer:
    def __init__(self, hotel_capacity, iva_rate=0.1):
        self.hotel_capacity = hotel_capacity
        self.iva_rate = iva_rate
        self.data = None
        self.group_request = None
        self.decision_params = None
        self.room_types = None
    
    def set_data(self, data_df):
        self.data = data_df
        return self
    
    def set_group_request(self, start_date, end_date, num_rooms, adr_lordo, adr_netto=None, fb_revenue=0, meeting_revenue=0, other_revenue=0):
        if adr_netto is None:
            adr_netto = adr_lordo / (1 + self.iva_rate)
            
        try:
            date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        except Exception as e:
            st.error(f"Errore nella creazione del date range: {e}")
            date_range = pd.date_range(start=datetime.now(), end=datetime.now() + timedelta(days=3))
            
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
    
    def set_group_request_variable(self, start_date, end_date, rooms_data, adr_lordo, adr_netto=None, fb_revenue=0, meeting_revenue=0, other_revenue=0):
        if adr_netto is None:
            adr_netto = adr_lordo / (1 + self.iva_rate)
            
        try:
            date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        except Exception as e:
            st.error(f"Errore nella creazione del date range: {e}")
            date_range = pd.date_range(start=datetime.now(), end=datetime.now() + timedelta(days=3))
            
        total_days = len(date_range)
        
        daily_fb = fb_revenue / total_days if total_days > 0 else 0
        daily_meeting = meeting_revenue / total_days if total_days > 0 else 0
        daily_other = other_revenue / total_days if total_days > 0 else 0
        total_ancillary = daily_fb + daily_meeting + daily_other
        
        base_df = pd.DataFrame({'data': date_range})
        rooms_data = pd.merge(base_df, rooms_data, on='data', how='left')
        rooms_data['camere'] = rooms_data['camere'].fillna(0)
        
        self.group_request = pd.DataFrame({
            'data': date_range,
            'camere_gruppo': rooms_data['camere'].values,
            'adr_gruppo_lordo': adr_lordo,
            'adr_gruppo_netto': adr_netto,
            'revenue_camere_gruppo': rooms_data['camere'].values * adr_netto,
            'revenue_fb_gruppo': daily_fb,
            'revenue_meeting_gruppo': daily_meeting,
            'revenue_other_gruppo': daily_other,
            'revenue_ancillare_gruppo': total_ancillary,
            'revenue_totale_gruppo': (rooms_data['camere'].values * adr_netto) + total_ancillary
        })
        
        return self
    
    def set_group_request_with_types(self, start_date, end_date, room_types, adr_lordo, adr_netto=None, fb_revenue=0, meeting_revenue=0, other_revenue=0):
        if adr_netto is None:
            adr_netto = adr_lordo / (1 + self.iva_rate)
            
        try:
            date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        except Exception as e:
            st.error(f"Errore nella creazione del date range: {e}")
            date_range = pd.date_range(start=datetime.now(), end=datetime.now() + timedelta(days=3))
            
        total_days = len(date_range)
        
        daily_fb = fb_revenue / total_days if total_days > 0 else 0
        daily_meeting = meeting_revenue / total_days if total_days > 0 else 0
        daily_other = other_revenue / total_days if total_days > 0 else 0
        total_ancillary = daily_fb + daily_meeting + daily_other
        
        if isinstance(room_types, pd.DataFrame):
            room_types = room_types.to_dict('records')
        elif not isinstance(room_types, list):
            st.error(f"Tipo di room_types non supportato: {type(room_types)}")
            room_types = [{"tipo": "ROH", "numero": 1, "adr_addon": 0.0}]
        
        total_rooms = sum(rt["numero"] for rt in room_types)
        total_revenue = sum((rt["numero"] * (adr_netto + rt["adr_addon"] / (1 + self.iva_rate))) for rt in room_types)
        
        if total_rooms > 0:
            weighted_adr_netto = total_revenue / total_rooms
        else:
            weighted_adr_netto = adr_netto
        
        weighted_adr_lordo = weighted_adr_netto * (1 + self.iva_rate)
        
        self.group_request = pd.DataFrame({
            'data': date_range,
            'camere_gruppo': total_rooms,
            'adr_gruppo_lordo': weighted_adr_lordo,
            'adr_gruppo_netto': weighted_adr_netto,
            'revenue_camere_gruppo': total_rooms * weighted_adr_netto,
            'revenue_fb_gruppo': daily_fb,
            'revenue_meeting_gruppo': daily_meeting,
            'revenue_other_gruppo': daily_other,
            'revenue_ancillare_gruppo': total_ancillary,
            'revenue_totale_gruppo': (total_rooms * weighted_adr_netto) + total_ancillary
        })
        
        self.room_types = room_types
        
        return self
    
    def set_decision_parameters(self, params):
        self.decision_params = params
        return self
    
    def analyze(self):
        if self.data is None or self.group_request is None or self.decision_params is None:
            raise ValueError("Dati, richiesta gruppo o parametri decisionali mancanti")
        
        result = pd.merge(self.data, self.group_request, on='data', how='right')
        
        result['finale_rn'] = result['fcst_ind_rn'] + result['otb_ind_rn'] + result['grp_otb_rn']
        result['finale_opz_rn'] = result['finale_rn']
        
        result['camere_disponibili'] = self.hotel_capacity - result['finale_rn']
        result['camere_displaced'] = np.where(
            self.hotel_capacity - (result['finale_rn'] + result['camere_gruppo']) > 0,
            0,
            (result['finale_rn'] + result['camere_gruppo']) - self.hotel_capacity
        )
        
        result['camere_gruppo_accettate'] = result['camere_gruppo']
        
        result['revenue_displaced'] = result['camere_displaced'] * result['finale_adr']
        
        result['revenue_camere_gruppo_effettivo'] = result['camere_gruppo'] * result['adr_gruppo_netto']
        
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
        total_group_rooms = analysis_df['camere_gruppo'].sum()
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
        
        accepted_rooms = analysis_df['camere_gruppo'].sum()
        displaced_rooms = analysis_df['camere_displaced'].sum()
        
        avg_occ_current = analysis_df['occupazione_attuale'].mean()
        avg_occ_with_group = analysis_df['occupazione_con_gruppo'].mean()
        
        should_accept = total_impact > 0
        
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
            'room_profit': total_group_room_revenue - total_displaced_revenue,
            'total_rev_profit': total_impact,
            'profit_per_room': (total_group_room_revenue - total_displaced_revenue) / accepted_rooms if accepted_rooms > 0 else 0,
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


st.title("Hotel Group Displacement Analyzer v0.9.5r5")
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
    enable_booking_parser = st.toggle("Parsing automatico richieste booking", value=True,
                                   help="Estrae automaticamente dati dalle richieste dell'ufficio booking")
    
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

if enable_booking_parser:
    with st.expander("💌 Analisi rapida richiesta booking", expanded=True):
        st.info("Incolla qui il testo della richiesta dell'ufficio booking per compilare automaticamente tutti i campi")
        booking_text = st.text_area("Testo richiesta", height=150, 
                                  placeholder="Esempio: Dal 25 marzo al 30 maggio 2025, nome agenzia: Example Tours, 25 camere")
        
        if booking_text and st.button("Analizza richiesta", key="parse_booking_main_btn"):
            with st.spinner("Analisi in corso..."):
                parsed_data = parse_booking_request(booking_text)
                
                if any(v is not None for v in parsed_data.values()):
                    st.success("Dati estratti dalla richiesta:")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if parsed_data['group_name']:
                            st.info(f"**Nome gruppo:** {parsed_data['group_name']}")
                        if parsed_data['num_rooms']:
                            st.info(f"**Numero camere:** {parsed_data['num_rooms']}")
                    
                    with col2:
                        if parsed_data['arrival_date']:
                            st.info(f"**Data arrivo:** {parsed_data['arrival_date'].strftime('%d/%m/%Y')}")
                        if parsed_data['departure_date']:
                            st.info(f"**Data partenza:** {parsed_data['departure_date'].strftime('%d/%m/%Y')}")
                            if not parsed_data['is_checkout']:
                                st.info("La data di partenza include il pernottamento (non è checkout)")
                    
                    if st.button("Conferma e utilizza questi dati", key="confirm_parsed_main_data"):
                        try:
                            import json
                            booking_data_json = json.dumps({
                                'group_name': parsed_data['group_name'],
                                'arrival_date': parsed_data['arrival_date'].strftime('%Y-%m-%d') if parsed_data['arrival_date'] else None,
                                'departure_date': parsed_data['departure_date'].strftime('%Y-%m-%d') if parsed_data['departure_date'] else None,
                                'num_rooms': parsed_data['num_rooms'],
                                'timestamp': int(time.time())
                            })
                            
                            st.session_state['booking_data_json'] = booking_data_json
                            st.session_state['group_name_stored'] = parsed_data['group_name']
                            st.session_state['num_rooms_stored'] = parsed_data['num_rooms']
                            st.session_state['arrival_date_stored'] = parsed_data['arrival_date']
                            st.session_state['departure_date_stored'] = parsed_data['departure_date']
                            
                            st.session_state['force_update'] = True
                            
                            st.success("Dati confermati! Aggiornamento in corso...")
                            st.info(f"Debug: Chiavi in session_state: {list(st.session_state.keys())}")
                            time.sleep(0.5)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Errore nel salvataggio dei dati: {e}")
                            st.code(traceback.format_exc())
                else:
                    st.warning("Non è stato possibile estrarre informazioni dalla richiesta. Verifica il formato del testo.")

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
                            start=idv_cy_data['data'].min(),
                            end=idv_cy_data['data'].max()
                        )
                        st.session_state['available_dates'] = available_dates
                        
                        st.success(f"File elaborati con successo! Range date disponibili: {available_dates.min().strftime('%d/%m/%Y')} - {available_dates.max().strftime('%d/%m/%Y')}")
                    else:
                        st.warning("Sono necessari almeno i file IDV anno corrente e anno precedente")
            
            if 'raw_excel_data' in st.session_state:
                st.success(f"File già elaborati. Seleziona il periodo da analizzare.")
                
                if st.button("📊 Debug: Mostra i dati importati"):
                    st.write("IDV CY Data:")
                    st.write(st.session_state['raw_excel_data']['idv_cy'].head())
                    st.write("IDV LY Data:")
                    st.write(st.session_state['raw_excel_data']['idv_ly'].head())
                
                available_dates = st.session_state['available_dates']
                
                col1, col2 = st.columns(2)
                with col1:
                    start_date_default = st.session_state.get('arrival_date_stored', available_dates.min())
                    if isinstance(start_date_default, str):
                        start_date_default = datetime.strptime(start_date_default, '%Y-%m-%d').date()
                    
                    start_date = st.date_input(
                        "Data inizio analisi", 
                        value=start_date_default,
                        min_value=available_dates.min(),
                        max_value=available_dates.max(),
                        key="start_date_input"
                    )
                
                with col2:
                    end_date_default = st.session_state.get('departure_date_stored', available_dates.min() + timedelta(days=3))
                    if isinstance(end_date_default, str):
                        end_date_default = datetime.strptime(end_date_default, '%Y-%m-%d').date()
                    
                    end_date = st.date_input(
                        "Data fine analisi", 
                        value=end_date_default,
                        min_value=available_dates.min(),
                        max_value=available_dates.max(),
                        key="end_date_input"
                    )
                
                if st.button("Carica dati per il periodo selezionato", type="primary"):
                    with st.spinner("Elaborazione dati in corso..."):
                        try:
                            start_datetime = pd.to_datetime(start_date)
                            end_datetime = pd.to_datetime(end_date)
                            
                            st.write(f"Debug: Range date selezionato: {start_datetime} - {end_datetime}")
                            date_range = pd.date_range(start=start_datetime, end=end_datetime)
                            
                            processed_data = process_imported_data(
                                st.session_state['raw_excel_data']['idv_cy'],
                                st.session_state['raw_excel_data']['idv_ly'],
                                st.session_state['raw_excel_data']['grp_otb'],
                                st.session_state['raw_excel_data']['grp_opz'],
                                date_range
                            )
                            
                            if processed_data is not None:
                                st.write("Debug: Dati elaborati:")
                                st.write(processed_data.head())
                                
                                st.session_state['analyzed_data'] = processed_data
                                st.session_state['selected_start_date'] = start_date
                                st.session_state['selected_end_date'] = end_date
                                st.success("Dati elaborati con successo!")
                                time.sleep(0.5)
                                st.rerun()
                        except Exception as e:
                            st.error(f"Errore durante il caricamento dei dati: {str(e)}")
                            st.code(traceback.format_exc())
            
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
                },
                use_container_width=True
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
                },
                use_container_width=True
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
                },
                use_container_width=True
            )
    
    with st.expander("Mappatura campi (avanzato)", expanded=False):
        st.warning("Queste impostazioni sono utilizzate solo se il riconoscimento automatico fallisce")
        date_column_name = st.text_input("Nome colonna data", "Giorno")
        rn_column_name = st.text_input("Nome colonna room nights", "Room nights")
        adr_column_name = st.text_input("Nome colonna ADR", "ADR Cam")
       
elif data_source == "Inserimento manuale":
    st.header("1️⃣ Periodo di Analisi")
    
    def on_start_date_change():
        st.session_state['start_date_fixed'] = st.session_state['start_date_input']
    
    def on_end_date_change():
        st.session_state['end_date_fixed'] = st.session_state['end_date_input']
    
    col1, col2 = st.columns(2)
    with col1:
        start_date_default = st.session_state.get('arrival_date_stored', datetime.now() + timedelta(days=30))
        if isinstance(start_date_default, str):
            start_date_default = datetime.strptime(start_date_default, '%Y-%m-%d').date()
        
        start_date = st.date_input("Data inizio analisi", 
                                 value=start_date_default, 
                                 key="start_date_input",
                                 on_change=on_start_date_change)
    
    with col2:
        end_date_default = st.session_state.get('departure_date_stored', datetime.now() + timedelta(days=33))
        if isinstance(end_date_default, str):
            end_date_default = datetime.strptime(end_date_default, '%Y-%m-%d').date()
        
        end_date = st.date_input("Data fine analisi", 
                               value=end_date_default, 
                               key="end_date_input",
                               on_change=on_end_date_change)
    
    st.header("2️⃣ Dati On The Books e Forecast")
    
    st.info("Inserisci manualmente i dati per il periodo selezionato")
    
    try:
        date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        
        if 'wizard_step' not in st.session_state and enable_wizard:
            st.session_state['wizard_step'] = 1
        
        if 'forecast_method' not in st.session_state:
            st.session_state['forecast_method'] = "LY - OTB"
            st.session_state['pickup_factor'] = 1.0
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
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
                    "giorno": st.column_config.TextColumn("Giorno", disabled=True),
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
                    "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
                    "giorno": st.column_config.TextColumn("Giorno", disabled=True),
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
                    "data_ly": st.column_config.DatetimeColumn("Data LY", format="DD/MM/YYYY", disabled=True),
                    "giorno_ly": st.column_config.TextColumn("Giorno LY", disabled=True),
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
                    "data_ly": st.column_config.DatetimeColumn("Data LY", format="DD/MM/YYYY", disabled=True),
                    "giorno_ly": st.column_config.TextColumn("Giorno LY", disabled=True),
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
                                                 ["LY - OTB", "Basato su LY", "Percentuale su OTB", "Valore assoluto"],
                                                 index=["LY - OTB", "Basato su LY", "Percentuale su OTB", "Valore assoluto"].index(st.session_state['forecast_method']))
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
                    elif forecast_method == "Valore assoluto":
                        pickup_value = st.number_input("Camere da aggiungere", 0, 100, st.session_state['pickup_value'],
                                                    help="Aggiunge questo numero di camere all'OTB attuale")
                        st.session_state['pickup_value'] = pickup_value
                    else:  # LY - OTB
                        st.info("Il forecast è calcolato come LY IND - OTB IND (pickup rimanente dall'anno precedente)")
            
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
            elif st.session_state['forecast_method'] == "Valore assoluto":
                final_data['fcst_ind_rn'] = final_data['otb_ind_rn'] + st.session_state['pickup_value']
            else:  # LY - OTB (default)
                final_data['fcst_ind_rn'] = final_data['ly_ind_rn'] - final_data['otb_ind_rn']
                final_data['fcst_ind_rn'] = final_data['fcst_ind_rn'].apply(lambda x: max(0, x))
           
            final_data['fcst_ind_adr'] = final_data['otb_ind_adr']
           
            final_data['otb_ind_rev'] = final_data['otb_ind_rn'] * final_data['otb_ind_adr']
            final_data['ly_ind_rev'] = final_data['ly_ind_rn'] * final_data['ly_ind_adr']
            final_data['grp_otb_rev'] = final_data['grp_otb_rn'] * final_data['grp_otb_adr']
            final_data['grp_opz_rev'] = final_data['grp_opz_rn'] * final_data['grp_opz_adr']
            final_data['fcst_ind_rev'] = final_data['fcst_ind_rn'] * final_data['fcst_ind_adr']
           
            final_data['finale_rn'] = final_data['fcst_ind_rn'] + final_data['otb_ind_rn'] + final_data['grp_otb_rn']
            final_data['finale_opz_rn'] = final_data['finale_rn'] + final_data['grp_opz_rn']
           
            final_data['finale_rev'] = final_data['otb_ind_rev'] + final_data['fcst_ind_rev'] + final_data['grp_otb_rev']
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
                        },
                        use_container_width=True
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
                        },
                        use_container_width=True
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
                        },
                        use_container_width=True
                    )
           
            analyzed_data = final_data
           
        except Exception as e:
            if not enable_wizard or st.session_state.get('wizard_step') == 6:
                st.error(f"Errore nel calcolo del forecast: {e}")
                st.code(traceback.format_exc())
            analyzed_data = None
    
    except Exception as e:
        st.error(f"Errore nell'elaborazione delle date: {e}")
        st.code(traceback.format_exc())
        analyzed_data = None

if start_date is None or end_date is None:
    st.header("1️⃣ Periodo di Analisi")
    
    col1, col2 = st.columns(2)
    with col1:
        default_start = st.session_state.get('arrival_date_stored', datetime.now() + timedelta(days=30))
        if isinstance(default_start, str):
            default_start = datetime.strptime(default_start, '%Y-%m-%d').date()
        start_date = st.date_input("Data inizio analisi", value=default_start, key="start_date_input")
    with col2:
        default_end = st.session_state.get('departure_date_stored', datetime.now() + timedelta(days=33))
        if isinstance(default_end, str):
            default_end = datetime.strptime(default_end, '%Y-%m-%d').date()
        end_date = st.date_input("Data fine analisi", value=default_end, key="end_date_input")
elif data_source == "Import file Excel" and 'analyzed_data' in st.session_state:
    st.header("1️⃣ Periodo di Analisi")
    st.info(f"Periodo di analisi: dal {start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')}")

st.header("3️⃣ Dettagli Richiesta Gruppo")

booking_data = get_booking_data()

col1, col2 = st.columns(2)
with col1:
    default_group_name = st.session_state.get('group_name_stored', booking_data.get('group_name', "Corporate Meeting") if booking_data else "Corporate Meeting")
    group_name = st.text_input("Nome Gruppo", value=default_group_name, key="group_name_input")
    
    st.subheader("Configurazione Camere")
    
    room_config_option = st.radio(
        "Configurazione camere",
        options=["Contingente fisso ROH", "Camere variabili per giorno", "Multiple tipologie"],
        index=0,
        help="Scegli come configurare l'offerta di camere per questo gruppo"
    )
    
    if room_config_option == "Contingente fisso ROH":
        default_num_rooms = st.session_state.get('num_rooms_stored', booking_data.get('num_rooms', 25) if booking_data else 25)
        num_rooms = st.number_input("Numero camere ROH", min_value=1, value=default_num_rooms, key="num_rooms_input")
        room_types = [{"tipo": "ROH", "numero": num_rooms, "adr_addon": 0.0}]
        
    elif room_config_option == "Camere variabili per giorno":
        st.info("Aggiungi dettagli specifici per giorno nella sezione sottostante")
        default_num_rooms = st.session_state.get('num_rooms_stored', booking_data.get('num_rooms', 25) if booking_data else 25)
        num_rooms = st.number_input("Numero camere medio", min_value=1, value=default_num_rooms, key="num_rooms_input")
        room_types = [{"tipo": "ROH", "numero": num_rooms, "adr_addon": 0.0}]
        
    elif room_config_option == "Multiple tipologie":
        st.subheader("Tipologie di camere")
        
        if 'room_types_df' not in st.session_state:
            default_num_rooms = st.session_state.get('num_rooms_stored', booking_data.get('num_rooms', 25) if booking_data else 25)
            st.session_state.room_types_df = pd.DataFrame({
                'tipo': ['ROH', 'Superior', 'Deluxe'],
                'numero': [max(default_num_rooms-10, 15), 8, 2],
                'adr_addon': [0.0, 30.0, 50.0]
            })
        
        edited_types = st.data_editor(
            st.session_state.room_types_df,
            hide_index=True,
            num_rows="dynamic",
            key="room_types_editor",
            column_config={
                "tipo": st.column_config.TextColumn("Tipologia"),
                "numero": st.column_config.NumberColumn("Numero", min_value=0, step=1, format="%d"),
                "adr_addon": st.column_config.NumberColumn("Supp. ADR (€)", min_value=0.0, format="€%.2f")
            },
            use_container_width=True
        )
        
        st.session_state.room_types_df = edited_types
        
        num_rooms = edited_types['numero'].sum()
        st.info(f"Totale: {num_rooms} camere")
        
        room_types = edited_types.to_dict('records')
    
    default_arrival = st.session_state.get('arrival_date_stored', booking_data.get('arrival_date', start_date) if booking_data else start_date)
    if isinstance(default_arrival, str):
        default_arrival = datetime.strptime(default_arrival, '%Y-%m-%d').date()
    group_arrival = st.date_input("Data di arrivo", value=default_arrival, key="arrival_date_input")
    
    default_departure = st.session_state.get('departure_date_stored', booking_data.get('departure_date', end_date) if booking_data else end_date)
    if isinstance(default_departure, str):
        default_departure = datetime.strptime(default_departure, '%Y-%m-%d').date()
    group_departure = st.date_input("Data di partenza", value=default_departure, key="departure_date_input")
    
with col2:
    adr_lordo = st.number_input("ADR base proposta (€ lordi)", min_value=0.0, value=900.0, key="adr_lordo_input")
    adr_netto = adr_lordo / (1 + iva_rate)
    st.info(f"ADR netto base: €{adr_netto:.2f}")
    
    if room_config_option == "Multiple tipologie":
        weighted_adr = 0
        total_rooms = sum(room_type["numero"] for room_type in room_types)
        
        if total_rooms > 0:
            weighted_adr = sum((room_type["numero"] / total_rooms) * (adr_lordo + room_type["adr_addon"]) 
                               for room_type in room_types)
            st.info(f"ADR medio lordo: €{weighted_adr:.2f}")
    
    fb_revenue = st.number_input("Revenue F&B previsto (€)", min_value=0.0, value=0.0, key="fb_revenue_input")
    meeting_revenue = st.number_input("Revenue sale riunioni (€)", min_value=0.0, value=0.0, key="meeting_revenue_input")
    other_revenue = st.number_input("Altro revenue ancillare (€)", min_value=0.0, value=0.0, key="other_revenue_input")
    
    total_ancillary = fb_revenue + meeting_revenue + other_revenue
    st.info(f"Totale revenue ancillare: €{total_ancillary:.2f}")
    
    try:
        date_nights = (group_departure - group_arrival).days
        
        if room_config_option == "Multiple tipologie":
            total_value = sum((room_type["numero"] * (adr_lordo + room_type["adr_addon"]) * date_nights) 
                              for room_type in room_types) + total_ancillary
        else:
            total_value = (adr_lordo * num_rooms * date_nights) + total_ancillary
        
        if total_value > 35000:
            st.warning(f"⚠️ Valore totale: €{total_value:,.2f} - Richiede autorizzazione (>€35.000)")
        else:
            st.success(f"✅ Valore totale: €{total_value:,.2f}")
    except Exception as e:
        st.error(f"Errore nel calcolo del valore totale: {e}")
        st.code(traceback.format_exc())

if room_config_option == "Camere variabili per giorno":
    st.subheader("Dettaglio camere per giorno")
    
    try:
        date_range = pd.date_range(start=group_arrival, end=group_departure - timedelta(days=1))
        
        if 'rooms_by_day_df' not in st.session_state:
            st.session_state.rooms_by_day_df = pd.DataFrame({
                'data': date_range,
                'giorno': [d.strftime('%a') for d in date_range],
                'camere': [num_rooms] * len(date_range)
            })
        elif len(st.session_state.rooms_by_day_df) != len(date_range) or not all(st.session_state.rooms_by_day_df['data'].isin(date_range)):
            st.session_state.rooms_by_day_df = pd.DataFrame({
                'data': date_range,
                'giorno': [d.strftime('%a') for d in date_range],
                'camere': [num_rooms] * len(date_range)
            })
        
        edited_rooms = st.data_editor(
            st.session_state.rooms_by_day_df,
            hide_index=True,
            key="variable_rooms_editor",
            column_config={
                "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
                "giorno": st.column_config.TextColumn("Giorno", disabled=True),
                "camere": st.column_config.NumberColumn("Camere", min_value=0, step=1, format="%d")
            },
            use_container_width=True
        )
        
        st.session_state.rooms_by_day_df = edited_rooms
        
        total_rooms = edited_rooms['camere'].sum()
        avg_rooms = edited_rooms['camere'].mean()
        st.info(f"Totale: {total_rooms} camere-notte, Media: {avg_rooms:.1f} camere/notte")
        
        num_rooms = int(avg_rooms)
    except Exception as e:
        st.error(f"Errore nella gestione delle camere variabili: {e}")
        st.code(traceback.format_exc())

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
    try:
        date_options = pd.date_range(start=group_arrival, end=group_departure - timedelta(days=1))
        formatted_date_options = [f"{d.strftime('%a')} {d.strftime('%d/%m/%Y')}" for d in date_options]
        date_dict = dict(zip(formatted_date_options, date_options))
    
        selected_formatted_dates = st.multiselect(
            "Seleziona date da includere nell'analisi (lascia vuoto per tutte)",
            options=formatted_date_options,
            default=formatted_date_options
        )
    
        dates_for_analysis = [date_dict[d] for d in selected_formatted_dates] if selected_formatted_dates else date_options
    except Exception as e:
        st.error(f"Errore nella preparazione delle date di analisi: {e}")
        st.code(traceback.format_exc())
        dates_for_analysis = []

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
                - Camere: {num_rooms} {"ROH" if room_config_option == "Contingente fisso ROH" else "di media" if room_config_option == "Camere variabili per giorno" else "in diverse tipologie"}
                - Giorni analizzati: {len(dates_for_analysis)} di {len(date_options)}
                """)
            
            with col2:
                st.info(f"""
                **Dettagli Economici**
                - ADR lordo: €{adr_lordo:.2f}{" (base)" if room_config_option == "Multiple tipologie" else ""}
                - ADR netto: €{adr_netto:.2f}{" (base)" if room_config_option == "Multiple tipologie" else ""}
                {f"- ADR medio: €{weighted_adr:.2f}" if room_config_option == "Multiple tipologie" else ""}
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
       
            if room_config_option == "Contingente fisso ROH":
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
            elif room_config_option == "Camere variabili per giorno":
                analyzer.set_group_request_variable(
                    start_date=group_arrival,
                    end_date=group_departure,
                    rooms_data=edited_rooms[['data', 'camere']],
                    adr_lordo=adr_lordo,
                    adr_netto=adr_netto,
                    fb_revenue=fb_revenue,
                    meeting_revenue=meeting_revenue,
                    other_revenue=other_revenue
                )
            elif room_config_option == "Multiple tipologie":
                room_types_list = edited_types.to_dict('records') if isinstance(edited_types, pd.DataFrame) else room_types
                
                analyzer.set_group_request_with_types(
                    start_date=group_arrival,
                    end_date=group_departure,
                    room_types=room_types_list,
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
                
                current_year = datetime.now().year
                is_future_year = group_arrival.year > current_year
                
                avg_adr_cy_ly = (metrics['avg_adr_cy'] + metrics['avg_adr_ly']) / 2
                avg_adr_label = "Media CY/LY"
                
                avg_occ = metrics['avg_occ_current']
                if avg_occ < 60:
                    future_increment = 0.03
                elif avg_occ < 80:
                    future_increment = 0.05
                else:
                    future_increment = 0.07
                
                future_adr = metrics['avg_adr_cy'] * (1 + future_increment)
                future_adr_label = f"+{future_increment*100:.0f}% (Anno Successivo)"
                
                for variation in adr_variations:
                    new_adr = adr_lordo * (1 + variation/100)
                    new_adr_netto = new_adr / (1 + iva_rate)
                    
                    analyzer_scenario = ExcelCompatibleDisplacementAnalyzer(hotel_capacity=hotel_capacity, iva_rate=iva_rate)
                    analyzer_scenario.set_data(analyzed_data)
                    analyzer_scenario.set_decision_parameters(decision_parameters)
                    
                    if room_config_option == "Contingente fisso ROH":
                        analyzer_scenario.set_group_request(
                            start_date=group_arrival,
                            end_date=group_departure,
                            num_rooms=num_rooms,
                            adr_lordo=new_adr,
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Camere variabili per giorno":
                        analyzer_scenario.set_group_request_variable(
                            start_date=group_arrival,
                            end_date=group_departure,
                            rooms_data=edited_rooms[['data', 'camere']],
                            adr_lordo=new_adr,
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Multiple tipologie":
                        scaled_room_types = []
                        for rt in room_types_list:
                            scaled_rt = rt.copy()
                            scaled_rt['adr_addon'] = rt['adr_addon'] * (1 + variation/100)
                            scaled_room_types.append(scaled_rt)
                            
                        analyzer_scenario.set_group_request_with_types(
                            start_date=group_arrival,
                            end_date=group_departure,
                            room_types=scaled_room_types,
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
                        "variation_label": f"{variation:+d}%",
                        "adr_lordo": new_adr,
                        "adr_netto": new_adr_netto,
                        "total_impact": metrics_scenario['total_impact'],
                        "room_profit": metrics_scenario['room_profit'],
                        "total_rev_profit": metrics_scenario['total_rev_profit'],
                        "displaced_rooms": metrics_scenario['displaced_rooms'],
                        "should_accept": metrics_scenario['should_accept'],
                        "total_lordo": metrics_scenario['total_lordo']
                    })
                
                if not is_future_year:
                    variation_perc = ((avg_adr_cy_ly / adr_netto) - 1) * 100
                    
                    analyzer_scenario = ExcelCompatibleDisplacementAnalyzer(hotel_capacity=hotel_capacity, iva_rate=iva_rate)
                    analyzer_scenario.set_data(analyzed_data)
                    analyzer_scenario.set_decision_parameters(decision_parameters)
                    
                    if room_config_option == "Contingente fisso ROH":
                        analyzer_scenario.set_group_request(
                            start_date=group_arrival,
                            end_date=group_departure,
                            num_rooms=num_rooms,
                            adr_lordo=avg_adr_cy_ly * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Camere variabili per giorno":
                        analyzer_scenario.set_group_request_variable(
                            start_date=group_arrival,
                            end_date=group_departure,
                            rooms_data=edited_rooms[['data', 'camere']],
                            adr_lordo=avg_adr_cy_ly * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Multiple tipologie":
                        factor = avg_adr_cy_ly / adr_netto
                        scaled_room_types = []
                        for rt in room_types_list:
                            scaled_rt = rt.copy()
                            scaled_rt['adr_addon'] = rt['adr_addon'] * factor
                            scaled_room_types.append(scaled_rt)
                            
                        analyzer_scenario.set_group_request_with_types(
                            start_date=group_arrival,
                            end_date=group_departure,
                            room_types=scaled_room_types,
                            adr_lordo=avg_adr_cy_ly * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    
                    result_scenario = analyzer_scenario.analyze()
                    if dates_for_analysis and len(dates_for_analysis) < len(date_options):
                        result_scenario = result_scenario[result_scenario['data'].isin(dates_for_analysis)]
                    
                    metrics_scenario = analyzer_scenario.get_summary_metrics(result_scenario)
                    
                    scenario_results.append({
                        "variation": variation_perc,
                        "variation_label": avg_adr_label,
                        "adr_lordo": avg_adr_cy_ly * (1 + iva_rate),
                        "adr_netto": avg_adr_cy_ly,
                        "total_impact": metrics_scenario['total_impact'],
                        "room_profit": metrics_scenario['room_profit'],
                        "total_rev_profit": metrics_scenario['total_rev_profit'],
                        "displaced_rooms": metrics_scenario['displaced_rooms'],
                        "should_accept": metrics_scenario['should_accept'],
                        "total_lordo": metrics_scenario['total_lordo']
                    })
                
                if is_future_year:
                    variation_perc = ((future_adr / adr_netto) - 1) * 100
                    
                    analyzer_scenario = ExcelCompatibleDisplacementAnalyzer(hotel_capacity=hotel_capacity, iva_rate=iva_rate)
                    analyzer_scenario.set_data(analyzed_data)
                    analyzer_scenario.set_decision_parameters(decision_parameters)
                    
                    if room_config_option == "Contingente fisso ROH":
                        analyzer_scenario.set_group_request(
                            start_date=group_arrival,
                            end_date=group_departure,
                            num_rooms=num_rooms,
                            adr_lordo=future_adr * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Camere variabili per giorno":
                        analyzer_scenario.set_group_request_variable(
                            start_date=group_arrival,
                            end_date=group_departure,
                            rooms_data=edited_rooms[['data', 'camere']],
                            adr_lordo=future_adr * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    elif room_config_option == "Multiple tipologie":
                        factor = future_adr / adr_netto
                        scaled_room_types = []
                        for rt in room_types_list:
                            scaled_rt = rt.copy()
                            scaled_rt['adr_addon'] = rt['adr_addon'] * factor
                            scaled_room_types.append(scaled_rt)
                            
                        analyzer_scenario.set_group_request_with_types(
                            start_date=group_arrival,
                            end_date=group_departure,
                            room_types=scaled_room_types,
                            adr_lordo=future_adr * (1 + iva_rate),
                            fb_revenue=fb_revenue,
                            meeting_revenue=meeting_revenue,
                            other_revenue=other_revenue
                        )
                    
                    result_scenario = analyzer_scenario.analyze()
                    if dates_for_analysis and len(dates_for_analysis) < len(date_options):
                        result_scenario = result_scenario[result_scenario['data'].isin(dates_for_analysis)]
                    
                    metrics_scenario = analyzer_scenario.get_summary_metrics(result_scenario)
                    
                    scenario_results.append({
                        "variation": variation_perc,
                        "variation_label": future_adr_label,
                        "adr_lordo": future_adr * (1 + iva_rate),
                        "adr_netto": future_adr,
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
                    "variation_label": "Scenario",
                    "variation": st.column_config.NumberColumn("Variazione %", format="%+.1f%%"),
                    "adr_lordo": st.column_config.NumberColumn("ADR Lordo", format="€%.2f"),
                    "adr_netto": st.column_config.NumberColumn("ADR Netto", format="€%.2f"),
                    "total_impact": st.column_config.NumberColumn("Impatto Rev", format="€%.2f"),
                    "room_profit": st.column_config.NumberColumn("Profitto Camere", format="€%.2f"),
                    "total_rev_profit": st.column_config.NumberColumn("Profitto Totale", format="€%.2f"),
                    "displaced_rooms": st.column_config.NumberColumn("Camere Displaced", format="%d"),
                    "should_accept": st.column_config.CheckboxColumn("Da Accettare")
                },
                use_container_width=True
            )
            
            scenarios_df_sorted = extended_analysis_results['scenarios_df'].sort_values('variation')
            
            fig_scenarios = px.line(
                scenarios_df_sorted, 
                x="variation_label", 
                y=["total_rev_profit", "room_profit"],
                markers=True,
                labels={
                    "variation_label": "Scenario ADR",
                    "value": "Profitto (€)",
                    "variable": "Tipo"
                },
                category_orders={"variation_label": scenarios_df_sorted["variation_label"].tolist()},
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
            if optimal_scenario["variation_label"] not in [f"{0:+d}%", "0%"]:
                if "Anno Successivo" in optimal_scenario["variation_label"] or "Media CY/LY" in optimal_scenario["variation_label"]:
                    st.info(f"""
                    💡 **Suggerimento tariffario ottimale**: 
                    L'ADR suggerita di €{optimal_scenario['adr_lordo']:.2f} ({optimal_scenario['variation_label']}) 
                    genererebbe un profitto totale di €{optimal_scenario['total_rev_profit']:,.2f}, 
                    con un incremento di €{optimal_scenario['total_rev_profit'] - metrics['total_rev_profit']:,.2f} 
                    rispetto all'ADR proposta originalmente.
                    
                    Questa tariffa è calcolata {"sulla media degli anni corrente e precedente" if "Media" in optimal_scenario["variation_label"] else "considerando un incremento basato sull'occupazione per l'anno successivo"}.
                    """)
                else:
                    st.info(f"""
                    💡 **Suggerimento tariffario ottimale**: 
                    Un'ADR di €{optimal_scenario['adr_lordo']:.2f} ({optimal_scenario['variation_label']}) 
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
                    },
                    use_container_width=True
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
               },
               use_container_width=True
           )
           
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(get_csv_download_link(result_df, f"displacement_{group_name}", "📥 Scarica dati completi (CSV)"), unsafe_allow_html=True)
        
        with col2:
            group_info = {
                'name': group_name,
                'arrival_date': group_arrival,
                'departure_date': group_departure,
                'num_rooms': num_rooms,
                'adr_lordo': adr_lordo,
                'adr_netto': adr_netto,
                'ancillary_revenue': total_ancillary
            }
            
            hotel_info = {
                'name': f"Hotel (capacità: {hotel_capacity} camere)",
                'capacity': hotel_capacity,
                'iva_rate': iva_rate
            }
            
            excel_download_link = get_excel_download_link(
                result_df, 
                metrics, 
                group_info, 
                hotel_info, 
                f"Report_Displacement_{group_name}_{datetime.now().strftime('%Y%m%d')}"
            )
            
            st.markdown(excel_download_link, unsafe_allow_html=True)
           
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
               
            nights = (group_departure - group_arrival).days
            
            email_text = generate_auth_email(
                group_name=group_name,
                total_revenue=metrics['total_lordo'],
                dates=result_df['data'].tolist(),
                rooms=num_rooms,
                adr=adr_lordo,
                nights=nights
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
                    
                    st.session_state['arrival_date_stored'] = next_start_date
                    st.session_state['departure_date_stored'] = next_end_date
                    st.rerun()
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
        },
        use_container_width=True
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
       <p>Hotel Group Displacement Analyzer | v0.9.5r5 developed by Alessandro Merella | Original excel concept and formulas by Andrea Conte<br>
       Sessione: {st.session_state['username']} | Ultimo accesso: {datetime.fromtimestamp(st.session_state['login_time']).strftime('%d/%m/%Y %H:%M')}<br>
       Distributed under MIT License
       </p>
   </div>
   """, 
   unsafe_allow_html=True
)
