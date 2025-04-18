# Changelog Hotel Group Displacement Analyzer

## v0.9.4r3 (Attuale)
- **Bugfix**: Risolto problema con il parsing automatico delle richieste booking
- **Miglioramento UX**: I dati vengono adesso correttamente caricati nei campi dopo l'analisi
- **Dipendenze**: Aggiunta dipendenza mancante requests nel requirements.txt

## v0.9.4r2
- **Miglioramento UI**: Corretti problemi di visualizzazione nel login
- **Bugfix**: Risolto problema con la visualizzazione del changelog
- **Miglioramento UX**: Il parser delle richieste ora compila anche le date di analisi

## v0.9.4
- **Nuova funzionalità**: Parsing automatico delle richieste di booking
- **Correzioni**: Risolti problemi di indentazione nel codice
- **Miglioramento UI**: Aggiunta validazione per i dati estratti automaticamente

## v0.9.3r3
- **Miglioramento UI**: Aggiunta la possibilità di configurare camere variabili per giorno
- **Miglioramento UI**: Aggiunto supporto per multiple tipologie di camera con supplementi
- **Miglioramento UX**: I dati inseriti sono ora più evidenti nell'interfaccia
- **Miglioramento Changelog**: Link "What's New" nella pagina di login invece di popup automatico
- **Correzioni**: Risolti problemi con parametri deprecati
- **Ottimizzazione**: Migliorata la navigazione tramite Tab nei data editor

## v0.9.2b2
- **Miglioramento Changelog**: Il changelog ora viene mostrato solo al primo accesso per ogni utente
- **Miglioramento Forecast**: Ripristinate tutte le modalità di calcolo del forecast con "LY - OTB" come opzione predefinita
- **Ottimizzazione**: Impostato il valore predefinito del moltiplicatore a 1.0

## v0.9.1
- **Correzione formula FCST IND**: Ora calcolata come LY IND - OTB IND
- **Nuova funzionalità**: Esportazione report in formato Excel formattato
- **Correzione display ADR**: Migliorata coerenza nei calcoli del valore totale gruppo
- **Miglioramento UI**: Corretti problemi di visualizzazione tabelle
- **Correzione bug**: Risolti problemi nell'importazione di file Excel
- **Modifiche funzionali**: L'analisi ora considera sempre tutte le camere del gruppo
- **Miglioria UX**: Implementato changelog al primo avvio

## v0.9.0beta2
- **Miglioramenti UI**: Tabelle con bordi colorati differenziati per anno corrente (verde) e anno precedente (arancione)
- **Correzioni bug**: Risolto problema di identificazione dei file Excel durante l'import
- **Nuova funzionalità**: Visualizzazione automatica del changelog al primo avvio
- **Correzioni UI**: Migliorata visualizzazione delle tabelle di inserimento dati manuale
- **Stabilità**: Migliorato il processo di autenticazione e gestione sessione
- **Integrazioni**: Supporto per changelog esterno in formato markdown

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
