# Galleon – Sistema report automatico

Tutto gira **sul PC del locale** (Python + database SQLite in un file). I dati grezzi non vanno su servizi esterni; l'AI serve solo per progettare e manutenere gli script.

## 1. Flusso

```
 Cassa (chiusura)  ─┐  export pianificato / email / cartella condivisa
 Zucchetti (SDI)   ─┼─►  dati/inbox/{cassa,fatture,presenze,inventari}
 In&Out (presenze) ─┘                  │
                         ogni notte alle 04:30 (Utilità di pianificazione)
                                       ▼
                     pipeline.py  →  db/galleon.sqlite  (storico completo)
                                       ▼
        report/Galleon_report.html  ·  report/Galleon_report.xlsx  ·  nota Obsidian "Galleon - Cruscotto"
```

I file importati finiscono in `dati/archivio`; quelli illeggibili in `dati/scarti` (motivo in `log/pipeline.log`). Se un file viene riesportato, i dati di quel giorno o documento vengono sostituiti, mai duplicati.

## 2. Cartelle

| Cartella | Contenuto | Chi la aggiorna |
|---|---|---|
| `dati/inbox/cassa` | `vendite_AAAA-MM-GG.csv`, `coperti_AAAA-MM-GG.csv` | gestionale cassa (automatico) |
| `dati/inbox/fatture` | fatture passive FatturaPA `.xml` / `.xml.p7m` | Zucchetti/SDI (automatico) |
| `dati/inbox/presenze` | `presenze_*.csv` | In&Out (automatico) |
| `dati/inbox/inventari` | `inventario_AAAA-MM-GG.csv` | **a mano, 1 volta al mese** |
| `anagrafiche/` | articoli, ricette, ingredienti, mappa acquisti, personale, costi fissi | a mano, solo quando cambia qualcosa |
| `report/` | report pronti + `storico/` | automatico |

## 3. Formato standard di input (separatore `;`, decimali `,`)

- **vendite**: `data;canale;codice;descrizione;categoria;quantita;importo_lordo;scontrinato(S/N)`
- **coperti**: `data;canale;coperti`
- **presenze**: `data;dipendente;ore;reparto`
- **inventario**: `ingrediente;quantita` (la data sta nel nome del file)
- **fatture**: XML FatturaPA standard. Va bene così com'è, senza conversioni.

Se gli export reali hanno un formato diverso, si aggiunge un piccolo *adattatore* che li converte in questo formato. I calcoli non cambiano.

## 4. Anagrafiche (il "dizionario" che collega i tre sistemi)

- `articoli.csv`: codice cassa → gruppo per il menu engineering (Pizze, Bevande, Cucina…), aliquota IVA
- `ricette.csv`: codice cassa → ingrediente + quantità per porzione (distinta base)
- `ingredienti.csv`: unità di misura + costo di riferimento (usato se mancano le fatture)
- `mappa_acquisti.csv`: P.IVA fornitore + testo contenuto nella descrizione della riga → ingrediente, **fattore** di conversione (es. sacco 25 kg → 25) e categoria di costo (`merci_food`, `merci_beverage`, `utenze`, `affitto`…)
- `personale.csv`: costo orario aziendale per dipendente (lordo + contributi + TFR + ratei)
- `costi_fissi.csv`: costi mensili che non arrivano da fattura

## 5. Frequenze

| Cosa | Quando |
|---|---|
| Import + rigenerazione di tutti i report | ogni notte (KPI della serata, 7 giorni, mese in corso) |
| Costo ingredienti | ricalcolato ogni notte (media degli acquisti degli ultimi 30 giorni) |
| Menu engineering | ogni notte, sugli ultimi 30 giorni |
| Conto economico definitivo e controllo sprechi | mensile, dopo l'inventario del giorno 1 |

## 6. Cosa contiene il report

1. **Cosa guardare oggi**: avvisi automatici (sprechi, food cost oltre obiettivo, rincari, utile/perdita)
2. **KPI**: ricavi netti, coperti, scontrino medio, ricavo per ora lavorata, costo lavoro %, food cost %
3. **Conto economico gestionale mensile**: ricavi → costo merci (con variazione rimanenze) → margine → personale → altri costi → risultato, con incidenze e **ricavi di pareggio**
4. **Menu engineering** (Kasavana-Smith, per gruppo): Star / Plowhorse / Puzzle / Dog, con l'azione consigliata
5. **Consumi e sprechi**: consumo reale (inventari + acquisti) confrontato con il consumo teorico (vendite × ricette), per ingrediente e in €
6. **Qualità dei dati**: righe di fattura non classificate, articoli senza ricetta, serate senza presenze, file scartati

## 7. Installazione su Windows

1. Installa Python 3.12 o superiore da python.org (spunta "Add to PATH"), poi esegui `pip install pandas openpyxl`
2. Copia la cartella in `C:\Galleon\galleon_report`
3. Prova: `python crea_dati_esempio.py`, poi `python pipeline.py`, poi apri `report\Galleon_report.html`
4. Per pianificarla, da prompt dei comandi:
   `schtasks /Create /SC DAILY /ST 04:30 /TN "Galleon Report" /TR "C:\Galleon\galleon_report\avvia_report.bat"`
5. Metti un collegamento a `report\Galleon_report.html` sul desktop: è sempre l'ultima versione.
6. Facoltativo: in `config.json` imposta `cartella_obsidian` sul vault "pietro" per avere la nota cruscotto aggiornata ogni notte.

> ⚠️ `crea_dati_esempio.py` azzera database e anagrafiche: va usato solo per la prova, mai con i dati veri.
