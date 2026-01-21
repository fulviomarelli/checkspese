# Checker Spese - Bot per pulizia dati Excel

Bot Python per la pulizia e validazione automatica dei dati delle spese da file Excel.

## Requisiti

- Python 3.7+
- openpyxl

## Installazione

```bash
# Attiva l'ambiente virtuale (se già creato)
source .env/bin/activate

# Installa le dipendenze
pip install -r requirements.txt
```

## Utilizzo

1. Posiziona il file Excel (.xlsx) da processare nella stessa directory del bot
2. Esegui il bot:

```bash
python checker_spese.py
```

3. Se ci sono più file .xlsx, il bot ti chiederà quale processare
4. Durante l'esecuzione potrebbero apparire dei modal per confermare correzioni

## Output

Il bot genera 3 file:

1. **clean_[nome_file].xlsx** - File pulito con i dati corretti
2. **modifiche_effettuate_[nome_file].txt** - Log dettagliato di tutte le modifiche
3. **errori.xlsx** - Righe con errori non risolvibili automaticamente (se presenti)

## Fasi del processo

### Fase 1: Eliminazione spese non POLIMI
Elimina tutte le righe dove il campo "Soggetto" non contiene "POLIMI"

### Fase 2: Filtraggio per stato
Elimina righe con stato diverso da:
- "Trasmessa"
- "Conclusa in attesa trasmissione attestazione"

### Fase 3: Eliminazione costi indiretti
Elimina righe con "Tipologia spesa" = "Costi indiretti"

### Fase 4: Pulizia dipartimenti
Corregge e valida i dipartimenti nel campo "Descrizione voce spesa":
- Correzione automatica errori comuni (spazi, typo, ecc.)
- Richiesta conferma per correzioni ambigue
- Segnalazione errori non correggibili

Dipartimenti validi:
- DAER, DCMC, DEIB, DENG, DICA, DIG_
- DMAT, DMEC, DASTU, DFIS, DESIGN, DABC

### Fase 5: Validazione rendicontazione
Verifica la coerenza tra:
- Tipologia spesa
- Inquadramento contrattuale
- Tipologia rendicontazione

Regole:
- Spese personale + Inquadramento valido → Costi standard
- Spese personale + Altro inquadramento → Costi reali
- Altre spese → Costi reali

## Note

- Il bot non modifica il file originale
- Tutti i cambiamenti vengono tracciati nel log
- Le righe problematiche vengono isolate nel file errori
- Interface grafica per conferme manuali quando necessario

## Creazione eseguibile Windows

Per creare un eseguibile .exe per Windows:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name CheckerSpese checker_spese.py
```

L'eseguibile sarà disponibile nella cartella `dist/`
