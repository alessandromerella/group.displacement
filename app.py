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

# -*- coding: utf-8 -*-

st.set_page_config(
    page_title="Hotel Group Displacement Analyzer v0.9.3", 
    layout="wide",
    menu_items=None
)

js_code = """
<script>
document.addEventListener('DOMContentLoaded', function() {
    // Imposta un breve ritardo per assicurarsi che i componenti Streamlit siano caricati
    setTimeout(function() {
        // Seleziona tutti gli input editabili nelle tabelle
        const editableCells = document.querySelectorAll('[data-testid="stDataEditorCell"] input, [data-testid="stDataEditorCell"] select');
        
        // Imposta i tabindex per gli input editabili
        editableCells.forEach((cell, index) => {
            cell.setAttribute('tabindex', index + 1);
        });
        
        // Gestisci l'evento keydown per il tasto Tab
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Tab') {
                const activeElement = document.activeElement;
                const isEditableCell = activeElement.closest('[data-testid="stDataEditorCell"]');
                
                if (isEditableCell) {
                    // Trova l'indice dell'elemento attivo
                    const allCells = Array.from(editableCells);
                    const currentIndex = allCells.indexOf(activeElement);
                    
                    // Calcola l'indice della cella successiva (o precedente se Shift+Tab)
                    let nextIndex;
                    if (e.shiftKey) {
                        nextIndex = currentIndex > 0 ? currentIndex - 1 : allCells.length - 1;
                    } else {
                        nextIndex = currentIndex < allCells.length - 1 ? currentIndex + 1 : 0;
                    }
                    
                    // Focalizza la cella successiva
                    if (allCells[nextIndex]) {
                        e.preventDefault();
                        allCells[nextIndex].focus();
                    }
                }
            }
        });
    }, 1000);
});
</script>
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
        
        st.image("https://raw.githubusercontent.com/alessandromerella/assets/main/logo_hotel_displacement.png", width=200)
        st.markdown("<h2 style='text-align: center;'>Group Displacement Analyzer</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Accedi per continuare</p>", unsafe_allow_html=True)
    
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login")
            
            if submit:
                try:
                    valid_usernames = st.secrets.credentials.usernames
                    valid_passwords = st.secrets.credentials.passwords
                except:
                    valid_usernames = ["revenue_manager", "general_manager", "sales_manager"]
                    valid_passwords = ["v2025", "vr2025", "2025"]
                    st.warning("Utilizzo credenziali di sviluppo. In produzione, configura i secrets.")
                
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
        
        st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
        st.markdown("<p style='font-size: 0.8rem; color: #777;'>v0.9.3</p>", unsafe_allow_html=True)
        
        if st.button("📋 What's New", key="whats_new_btn"):
            st.session_state['show_changelog'] = True
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    if 'show_changelog' in st.session_state and st.session_state['show_changelog']:
        with st.expander("Changelog - Novità della versione 0.9.3", expanded=True):
            st.markdown("""
            ### Changelog v0.9.3
            
            **Nuove funzionalità:**
            - 🔄 Aggiunto supporto per camere variabili per giorno
            - 🏨 Supporto per multiple tipologie di camera con supplementi ADR
            - 🔍 Migliorata l'interfaccia di selezione delle date per l'analisi
            - 📊 Data entry avanzato con valori evidenziati in grassetto
            - ⌨️ Supporto per navigazione con Tab tra tabelle (sperimentale)
            
            **Miglioramenti:**
            - 🖼️ Ridimensionato il logo nella pagina di login
            - 📱 Interfaccia ottimizzata per dispositivi mobili
            - 🔒 Migliorata la gestione delle sessioni
            
            **Correzioni:**
            - 🐛 Corretto il problema con la copia dell'email di autorizzazione
            - 🧮 Migliorati i calcoli di displacement con contingenti variabili
            - 🚀 Ottimizzate le prestazioni dell'applicazione
            """)
            
            if st.button("Chiudi", key="close_changelog"):
                st.session_state['show_changelog'] = False
                st.rerun()
    
    return False

if not authenticate():
    st.stop()

st.markdown(js_code, unsafe_allow_html=True)

st.sidebar.info(f"Accesso effettuato come: {st.session_state['username']}")
if st.sidebar.button("Logout"):
    for key in ['authenticated', 'username', 'login_time']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

COLOR_PALETTE = {
    "primary": "#D8C0B7",
    "secondary": "#8CA68C",
    "text": "#5E5E5E",
    "background": "#F8F6F4",
    "accent": "#B6805C",
    "positive": "#8CA68C",
    "negative": "#D8837F"
}

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
    
    /* Stile per far risaltare i dati inseriti manualmente */
    .stDataEditor input[type="number"],
    .stDataEditor input[type="text"] {{
        font-weight: bold !important;
        font-size: 1.05rem !important;
    }}
    
    /* Colore di sfondo per le celle modificabili */
    .stDataEditor div[data-testid="column"] {{
        background-color: rgba(220, 220, 220, 0.1) !important;
    }}
    
    /* Per celle modificate */
    .stDataEditor div[data-testid="cell"]:focus-within {{
        background-color: rgba(140, 166, 140, 0.15) !important;
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
        self.room_types = None
    
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
    
    def set_group_request_variable(self, start_date, end_date, rooms_data, adr_lordo, adr_netto=None, fb_revenue=0, meeting_revenue=0, other_revenue=0):
        if adr_netto is None:
            adr_netto = adr_lordo / (1 + self.iva_rate)
            
        date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        total_days = len(date_range)
        
        daily_fb = fb_revenue / total_days if total_days > 0 else 0
        daily_meeting = meeting_revenue / total_days if total_days > 0 else 0
        daily_other = other_revenue / total_days if total_days > 0 else 0
        total_ancillary = daily_fb + daily_meeting + daily_other
        
        # Prepare base DataFrame with all dates in range
        base_df = pd.DataFrame({'data': date_range})
        
        # Merge with provided room data
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
            
        date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))
        total_days = len(date_range)
        
        daily_fb = fb_revenue / total_days if total_days > 0 else 0
        daily_meeting = meeting_revenue / total_days if total_days > 0 else 0
        daily_other = other_revenue / total_days if total_days > 0 else 0
        total_ancillary = daily_fb + daily_meeting + daily_other
        
        # Calcola il numero totale di camere e ADR medio ponderato
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
        
        # Conserva le informazioni sulle tipologie per eventuali report dettagliati
        self.room_types = room_types
        
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
            'should_accept': total_rev_profit > 0,
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
    
    def create_visualizations(self, analysis_df, metrics):
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


st.title("Hotel Group Displacement Analyzer")
st.markdown("*Strumento di analisi richieste preventivo gruppi*")

with st.sidebar:
    st.header("Configurazione Hotel")
    hotel_capacity = st.number_input("Capacità hotel (camere)", min_value=1, value=66)
    iva_rate = st.number_input("Aliquota IVA (%)", min_value=0.0, max_value=30.0, value=10.0) / 100
    
    st.header("Fonte dati")
    data_source = st.radio("Seleziona fonte dati", ["Import file Excel", "Inserimento manuale"])
    
    if data_source == "Import file Excel":
        with st.expander("Configurazione file", expanded=True):
            st.info("Carica i file Excel esportati da Power BI per il periodo desiderato")
            
            idv_cy_file = st.file_uploader("File IDV Anno Corrente (OTB)", type=["xlsx", "xls"])
            idv_ly_file = st.file_uploader("File IDV Anno Precedente (LY)", type=["xlsx", "xls"])
            grp_otb_file = st.file_uploader("File Gruppi Confermati (OTB)", type=["xlsx", "xls"])
            grp_opz_file = st.file_uploader("File Gruppi Opzionati (OTB)", type=["xlsx", "xls"])
        
        with st.expander("Mappatura campi (avanzato)", expanded=False):
            st.warning("Da configurare in base ai nomi delle colonne nei tuoi file Excel")
            date_column_name = st.text_input("Nome colonna data", "Data")
            rn_column_name = st.text_input("Nome colonna room nights", "RN")
            adr_column_name = st.text_input("Nome colonna ADR", "ADR")

st.header("1️⃣ Periodo di Analisi")

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Data inizio analisi", value=datetime.now() + timedelta(days=30))
with col2:
    end_date = st.date_input("Data fine analisi", value=datetime.now() + timedelta(days=33))

st.header("2️⃣ Dati On The Books e Forecast")

date_range = pd.date_range(start=start_date, end=end_date - timedelta(days=1))

analyzed_data = None

if data_source == "Import file Excel":
    if idv_cy_file is not None and idv_ly_file is not None:
        with st.spinner("Elaborazione file in corso..."):
            processed_data = process_uploaded_files(
                idv_cy_file,
                idv_ly_file,
                grp_otb_file,
                grp_opz_file,
                date_range,
                date_column_name,
                rn_column_name,
                adr_column_name
            )
            
            if processed_data is not None:
                st.success("File elaborati con successo")
                
                st.subheader("Dati elaborati")
                tab1, tab2, tab3 = st.tabs(["Room Nights", "ADR", "Revenue"])
                
                with tab1:
                    st.dataframe(
                        processed_data[['data', 'giorno', 'otb_ind_rn', 'ly_ind_rn', 'fcst_ind_rn', 'grp_otb_rn', 'grp_opz_rn', 'finale_rn']],
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
                        processed_data[['data', 'giorno', 'otb_ind_adr', 'ly_ind_adr', 'fcst_ind_adr', 'grp_otb_adr', 'grp_opz_adr', 'finale_adr']],
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
                        processed_data[['data', 'giorno', 'otb_ind_rev', 'ly_ind_rev', 'fcst_ind_rev', 'grp_otb_rev', 'grp_opz_rev', 'finale_rev']],
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
               
                analyzed_data = processed_data
            else:
                st.error("Errore nell'elaborazione dei file")
                analyzed_data = None
    else:
        st.warning("Carica i file IDV per iniziare l'analisi")
        analyzed_data = None
       
elif data_source == "Inserimento manuale":
   st.info("Inserisci manualmente i dati per il periodo selezionato")
   
   # Prepara struttura base dei dati
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
   
   # Room Nights
   st.subheader("Inserimento Room Nights")
   col1, col2 = st.columns(2)
   
   with col1:
       st.markdown("**Room Nights - Anno Corrente**")
       edited_rn_cy = st.data_editor(
           df_base[['data', 'giorno', 'otb_ind_rn', 'grp_otb_rn', 'grp_opz_rn']],
           hide_index=True,
           key="rn_cy",
           disabled=False,
           column_config={
               "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
               "giorno": st.column_config.TextColumn("Giorno", disabled=True),
               "otb_ind_rn": st.column_config.NumberColumn("OTB IND", min_value=0, format="%d"),
               "grp_otb_rn": st.column_config.NumberColumn("GRP OTB", min_value=0, format="%d"),
               "grp_opz_rn": st.column_config.NumberColumn("GRP OPZ", min_value=0, format="%d")
           }
       )
   
   with col2:
       st.markdown("**Room Nights - Anno Precedente**")
       edited_rn_ly = st.data_editor(
           df_base[['data_ly', 'giorno_ly', 'ly_ind_rn']],
           hide_index=True,
           key="rn_ly",
           disabled=False,
           column_config={
               "data_ly": st.column_config.DateColumn("Data LY", format="DD/MM/YYYY", disabled=True),
               "giorno_ly": st.column_config.TextColumn("Giorno LY", disabled=True),
               "ly_ind_rn": st.column_config.NumberColumn("LY IND", min_value=0, format="%d")
           }
       )
   
   # ADR
   st.subheader("Inserimento ADR")
   col1, col2 = st.columns(2)
   
   with col1:
       st.markdown("**ADR - Anno Corrente**")
       edited_adr_cy = st.data_editor(
           df_base[['data', 'giorno', 'otb_ind_adr', 'grp_otb_adr', 'grp_opz_adr']],
           hide_index=True,
           key="adr_cy",
           disabled=False,
           column_config={
               "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
               "giorno": st.column_config.TextColumn("Giorno", disabled=True),
               "otb_ind_adr": st.column_config.NumberColumn("OTB IND", min_value=0, format="€%.2f"),
               "grp_otb_adr": st.column_config.NumberColumn("GRP OTB", min_value=0, format="€%.2f"),
               "grp_opz_adr": st.column_config.NumberColumn("GRP OPZ", min_value=0, format="€%.2f")
           }
       )
   
   with col2:
       st.markdown("**ADR - Anno Precedente**")
       edited_adr_ly = st.data_editor(
           df_base[['data_ly', 'giorno_ly', 'ly_ind_adr']],
           hide_index=True,
           key="adr_ly",
           disabled=False,
           column_config={
               "data_ly": st.column_config.DateColumn("Data LY", format="DD/MM/YYYY", disabled=True),
               "giorno_ly": st.column_config.TextColumn("Giorno LY", disabled=True),
               "ly_ind_adr": st.column_config.NumberColumn("LY IND", min_value=0, format="€%.2f")
           }
       )
   
   # Calcolo automatico del forecast
   st.subheader("Parametri Forecast")
   col1, col2 = st.columns(2)
   
   with col1:
       forecast_method = st.selectbox("Metodo di Forecast", 
                                     ["Basato su LY", "Percentuale su OTB", "Valore assoluto"])
   
   with col2:
       if forecast_method == "Basato su LY":
           pickup_factor = st.slider("Moltiplicatore LY", 0.5, 2.0, 1.1, 0.1, 
                                   help="Moltiplica i dati LY per questo fattore")
       elif forecast_method == "Percentuale su OTB":
           pickup_percentage = st.slider("Pickup %", 0, 100, 20, 5, 
                                       help="Aggiunge questa percentuale all'OTB attuale")
       else:
           pickup_value = st.number_input("Camere da aggiungere", 0, 100, 10,
                                        help="Aggiunge questo numero di camere all'OTB attuale")
   
   # Combinare i dati e calcolare il forecast
   try:
       # Combina i dati di room nights
       final_data = edited_rn_cy.copy()
       final_data = pd.merge(final_data, edited_rn_ly[['data_ly', 'ly_ind_rn']], 
                           left_index=True, right_index=True)
       
       # Combina i dati di ADR
       final_data = pd.merge(final_data, edited_adr_cy[['otb_ind_adr', 'grp_otb_adr', 'grp_opz_adr']], 
                           left_index=True, right_index=True)
       final_data = pd.merge(final_data, edited_adr_ly[['ly_ind_adr']], 
                           left_index=True, right_index=True)
       
       # Calcola il forecast basato sul metodo selezionato
       if forecast_method == "Basato su LY":
           final_data['fcst_ind_rn'] = np.ceil(final_data['ly_ind_rn'] * pickup_factor)
       elif forecast_method == "Percentuale su OTB":
           final_data['fcst_ind_rn'] = np.ceil(final_data['otb_ind_rn'] * (1 + pickup_percentage/100))
       else:
           final_data['fcst_ind_rn'] = final_data['otb_ind_rn'] + pickup_value
       
       # Calcola ADR del forecast (media ponderata tra OTB e nuovo business)
       final_data['fcst_ind_adr'] = final_data['otb_ind_adr']  # Semplificato per ora
       
       # Calcola totali
       final_data['finale_rn'] = final_data['fcst_ind_rn'] + final_data['grp_otb_rn']
       final_data['finale_opz_rn'] = final_data['fcst_ind_rn'] + final_data['grp_otb_rn'] + final_data['grp_opz_rn']
       
       # Calcola revenue
       final_data['otb_ind_rev'] = final_data['otb_ind_rn'] * final_data['otb_ind_adr']
       final_data['ly_ind_rev'] = final_data['ly_ind_rn'] * final_data['ly_ind_adr']
       final_data['grp_otb_rev'] = final_data['grp_otb_rn'] * final_data['grp_otb_adr']
       final_data['grp_opz_rev'] = final_data['grp_opz_rn'] * final_data['grp_opz_adr']
       final_data['fcst_ind_rev'] = final_data['fcst_ind_rn'] * final_data['fcst_ind_adr']
       
       final_data['finale_rev'] = final_data['fcst_ind_rev'] + final_data['grp_otb_rev']
       final_data['finale_adr'] = np.where(final_data['finale_rn'] > 0,
                                        final_data['finale_rev'] / final_data['finale_rn'],
                                        0)
       
       # Mostra il forecast calcolato
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
       st.error(f"Errore nel calcolo del forecast: {e}")
       analyzed_data = None

st.header("3️⃣ Dettagli Richiesta Gruppo")

col1, col2 = st.columns(2)
with col1:
    group_name = st.text_input("Nome Gruppo", "Corporate Meeting")
    
    st.subheader("Configurazione Camere")
    
    # Aggiungi opzioni per la gestione delle camere
    room_config_option = st.radio(
        "Configurazione camere",
        options=["Contingente fisso ROH", "Camere variabili per giorno", "Multiple tipologie"],
        index=0,
        help="Scegli come configurare l'offerta di camere per questo gruppo"
    )
    
    if room_config_option == "Contingente fisso ROH":
        # Visualizzazione standard per numero di camere fisso
        num_rooms = st.number_input("Numero camere ROH", min_value=1, value=25)
        room_types = [{"tipo": "ROH", "numero": num_rooms, "adr_addon": 0.0}]
        
    elif room_config_option == "Camere variabili per giorno":
        # Per ora inserisci un placeholder, implementeremo l'editor più avanti
        st.info("Aggiungi dettagli specifici per giorno nella sezione sottostante")
        num_rooms = st.number_input("Numero camere medio", min_value=1, value=25)
        room_types = [{"tipo": "ROH", "numero": num_rooms, "adr_addon": 0.0}]
        
    elif room_config_option == "Multiple tipologie":
        st.subheader("Tipologie di camere")
        
        # Creiamo un dataframe per le tipologie di camere
        if 'room_types_df' not in st.session_state:
            st.session_state.room_types_df = pd.DataFrame({
                'tipo': ['ROH', 'Superior', 'Deluxe'],
                'numero': [15, 8, 2],
                'adr_addon': [0.0, 30.0, 50.0]  # Supplementi rispetto all'ADR base
            })
        
        # Editor per le tipologie di camere
        edited_types = st.data_editor(
            st.session_state.room_types_df,
            hide_index=True,
            num_rows="dynamic",
            key="room_types_editor",
            column_config={
                "tipo": st.column_config.TextColumn("Tipologia"),
                "numero": st.column_config.NumberColumn("Numero", min_value=0, step=1, format="%d"),
                "adr_addon": st.column_config.NumberColumn("Supp. ADR (€)", min_value=0.0, format="€%.2f")
            }
        )
        
        # Aggiorna il dataframe nello session state
        st.session_state.room_types_df = edited_types
        
        # Calcola il totale delle camere
        num_rooms = edited_types['numero'].sum()
        st.info(f"Totale: {num_rooms} camere")
        
        # Crea la lista delle tipologie per l'analisi
        room_types = edited_types.to_dict('records')
    
    group_arrival = st.date_input("Data di arrivo", value=start_date, key="group_arrival")
    group_departure = st.date_input("Data di partenza", value=end_date, key="group_departure")
    
with col2:
    adr_lordo = st.number_input("ADR base proposta (€ lordi)", min_value=0.0, value=900.0)
    adr_netto = adr_lordo / (1 + iva_rate)
    st.info(f"ADR netto base: €{adr_netto:.2f}")
    
    # Calcola e mostra ADR medio se ci sono più tipologie
    if room_config_option == "Multiple tipologie":
        weighted_adr = 0
        total_rooms = sum(room_type["numero"] for room_type in room_types)
        
        if total_rooms > 0:
            weighted_adr = sum((room_type["numero"] / total_rooms) * (adr_lordo + room_type["adr_addon"]) 
                               for room_type in room_types)
            st.info(f"ADR medio lordo: €{weighted_adr:.2f}")
    
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

# Se l'opzione camere variabili è selezionata, mostra un editor per inserire il numero di camere per giorno
if room_config_option == "Camere variabili per giorno":
    st.subheader("Dettaglio camere per giorno")
    
    # Crea un dataframe con le date del soggiorno
    date_range = pd.date_range(start=group_arrival, end=group_departure - timedelta(days=1))
    rooms_by_day_df = pd.DataFrame({
        'data': date_range,
        'giorno': [d.strftime('%a') for d in date_range],
        'camere': [num_rooms] * len(date_range)  # Valore predefinito uguale a num_rooms
    })
    
    # Editor per camere variabili
    edited_rooms = st.data_editor(
        rooms_by_day_df,
        hide_index=True,
        key="variable_rooms_editor",
        column_config={
            "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY", disabled=True),
            "giorno": st.column_config.TextColumn("Giorno", disabled=True),
            "camere": st.column_config.NumberColumn("Camere", min_value=0, step=1, format="%d")
        }
    )
    
    total_rooms = edited_rooms['camere'].sum()
    avg_rooms = edited_rooms['camere'].mean()
    st.info(f"Totale: {total_rooms} camere-notte, Media: {avg_rooms:.1f} camere/notte")
    
    num_rooms = int(avg_rooms)

date_options = pd.date_range(start=group_arrival, end=group_departure - timedelta(days=1))
formatted_date_options = [f"{d.strftime('%a')} {d.strftime('%d/%m/%Y')}" for d in date_options]
date_dict = dict(zip(formatted_date_options, date_options))

selected_formatted_dates = st.multiselect(
    "Seleziona date da includere nell'analisi (lascia vuoto per tutte)",
    options=formatted_date_options,
    default=formatted_date_options
)

# Converti le date formattate selezionate in oggetti datetime
dates_for_analysis = [date_dict[d] for d in selected_formatted_dates] if selected_formatted_dates else date_options

st.header("4️⃣ Analisi Displacement")

if st.button("Esegui Analisi", type="primary", use_container_width=True):
    if analyzed_data is not None:
        # Popup di double check
        with st.expander("Verifica parametri analisi", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"""
                **Dettagli Gruppo**
                - Nome: {group_name}
                - Periodo: {group_arrival.strftime('%d/%m/%Y')} - {group_departure.strftime('%d/%m/%Y')} ({date_nights} notti)
                - Camere: {num_rooms} {'camere variabili' if room_config_option == 'Camere variabili per giorno' else 'ROH' if room_config_option == 'Contingente fisso ROH' else 'in diverse tipologie'}
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
            
        # Bottone di conferma
        confirm = st.button("Conferma e Procedi", key="confirm_analysis")
        
        if confirm:
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
                
                # Nel metodo di analisi, modifica la chiamata a set_group_request
                if room_config_option == "Contingente fisso ROH":
                    # Usa il numero fisso di camere
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
                    # Usa i dati delle camere variabili
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
                    # Usa le informazioni sulle tipologie di camere
                    analyzer.set_group_request_with_types(
                        start_date=group_arrival,
                        end_date=group_departure,
                        room_types=room_types,
                        adr_lordo=adr_lordo,
                        adr_netto=adr_netto,
                        fb_revenue=fb_revenue,
                        meeting_revenue=meeting_revenue,
                        other_revenue=other_revenue
                    )
                
                result_df = analyzer.analyze()
                
                # Filtra per date selezionate
                if dates_for_analysis and len(dates_for_analysis) < len(date_options):
                    result_df = result_df[result_df['data'].isin(dates_for_analysis)]
                
                metrics = analyzer.get_summary_metrics(result_df)
                
                detail_fig, summary_fig = analyzer.create_visualizations(result_df, metrics)
                
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
                
                st.subheader("Riepilogo Finanziario")
                
                financial_df = pd.DataFrame({
                    'Voce': ['TOT. LORDO', 'TOT. NETTO', 'Offerta', 'ADR netto', 'Ancillary', 'Room Profit', 
                           'Extra IND LY', 'Extra per room IND LY', 'Extra TY', 'Total Rev Profit'],
                    'Valore': [
                        f"€{metrics['total_lordo']:,.2f}",
                        f"€{metrics['group_room_revenue'] + metrics['group_ancillary']:,.2f}",
                        f"€{adr_lordo:.2f}",
                        f"€{adr_netto:.2f}",
                        f"€{metrics['group_ancillary']:,.2f}",
                        f"€{metrics['room_profit']:,.2f}",
                        f"€{(metrics['extra_vs_ly'] * metrics['accepted_rooms']):,.2f}",
                        f"€{metrics['extra_vs_ly']:,.2f}",
                        f"€{metrics['room_profit']:,.2f}",
                        f"€{metrics['total_rev_profit']:,.2f}"
                    ]
                })
                
                st.table(financial_df)
                
                st.subheader("Dati Dettagliati")
                display_cols = ['data', 'giorno', 'finale_rn', 'camere_gruppo', 'camere_disponibili', 
                              'camere_displaced', 'adr_gruppo_netto', 'finale_adr', 
                              'revenue_camere_gruppo_effettivo', 'revenue_displaced', 'impatto_revenue_totale']
                
                st.dataframe(
                    result_df[display_cols],
                    column_config={
                        "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                        "giorno": "Giorno",
                        "finale_rn": st.column_config.NumberColumn("OTB", format="%d"),
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
                
                if metrics['needs_authorization']:
                    st.warning("⚠️ Questa richiesta gruppo supera il valore di €35.000 e richiede autorizzazione")
                    
                    st.subheader("Email di Richiesta Autorizzazione")
                    
                    email_text = generate_auth_email(
                        group_name=group_name,
                        total_revenue=metrics['total_lordo'],
                        dates=result_df['data'].tolist(),
                        rooms=num_rooms,
                        adr=adr_lordo
                    )
                    
                    st.text_area("Email da inviare", email_text, height=300)
                    
                    email_text_base64 = base64.b64encode(email_text.encode('utf-8')).decode('utf-8')

                    st.markdown(
                        f"""
                        <button onclick="
                            const emailText = atob('{email_text_base64}');
                            navigator.clipboard.writeText(emailText);
                            alert('Email copiata negli appunti');
                        " style="
                            background-color: {COLOR_PALETTE['secondary']};
                            color: white;
                            border: none;
                            padding: 10px 20px;
                            border-radius: 5px;
                            cursor: pointer;
                            font-family: 'Inter', sans-serif;
                        ">📋 Copia Email</button>
                        """,
                        unsafe_allow_html=True
                    )
    else:
        st.error("Nessun dato disponibile per l'analisi. Assicurati di caricare i file necessari o di inserire i dati manualmente.")

st.markdown("---")
st.markdown(
    f"""
    <div style='text-align: center; font-family: Inter, sans-serif; color: #5E5E5E; font-size: 0.8rem;'>
        <p>Hotel Group Displacement Analyzer | v0.9.3 developed by Alessandro Merella | Original excel concept and formulas by Andrea Conte<br>
        Sessione: {st.session_state['username']} | Ultimo accesso: {datetime.fromtimestamp(st.session_state['login_time']).strftime('%d/%m/%Y %H:%M')}
        </p>
    </div>
    """, 
    unsafe_allow_html=True
)
