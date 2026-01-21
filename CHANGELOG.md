# Changelog

Tutte le modifiche importanti a questo progetto verranno documentate in questo file

## [Non rilasciato]

### In sviluppo
- Nessuna modifica in corso

---

## [1.0.0] - 2026-01-21

###  Prima Release

#### Aggiunto
- **Fase 1**: Eliminazione automatica spese non POLIMI
- **Fase 2**: Filtraggio per stati validi (Trasmessa, Conclusa in attesa)
- **Fase 3**: Eliminazione costi indiretti
- **Fase 4**: Pulizia intelligente dipartimenti
  - Correzione automatica errori di battitura
  - Rimozione spazi e prefissi errati
  - Modal interattivo per conferme manuali
- **Fase 5**: Validazione regole di rendicontazione
  - Controllo coerenza tra tipo spesa, inquadramento e rendicontazione
  - Modal con errori di validazione
- **Output**:
  - File Excel pulito (`clean_[nome].xlsx`)
  - Log dettagliato modifiche (`modifiche_effettuate_[nome].txt`)
  - File errori separato (`errori.xlsx`)
- **Launcher Windows**: Script BAT con controlli automatici
- **Documentazione completa**: README e ISTRUZIONI_UTENTE
- **Eseguibile standalone**: Nessuna dipendenza richiesta

#### Tecnologie
- Python 3.7+
- openpyxl per manipolazione Excel
- tkinter per interfaccia grafica
- PyInstaller per creazione eseguibile


---

[Non rilasciato]: https://github.com/TUO_USERNAME/checker_lavoro/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/TUO_USERNAME/checker_lavoro/releases/tag/v1.0.0
