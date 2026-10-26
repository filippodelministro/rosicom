# rosicom

Strumento locale e minimale per seguire la tua lega di [Leghe Fantacalcio](https://leghe.fantacalcio.it).
Scarica rose e calendario della lega e li mostra in una pagina nel browser:

- **Squadre**: la rosa di ogni squadra, divisa per ruolo
- **Classifica**: punti, vittorie/pareggi/sconfitte, gol fatti e subiti, fantapunti
- **Calendario**: le partite giocate, con risultato e fantapunti, e quelle da giocare

Tutto gira sul tuo computer: nessun server, nessun dato inviato altrove.

## Requisiti

- Python 3.10 o superiore
- Un account Leghe Fantacalcio che partecipa alla lega

```bash
git clone https://github.com/filippodelministro/rosicom.git
cd rosicom
pip install -r requirements.txt
```

## Provalo subito con una lega inventata

```bash
python scripts/genera_esempio.py      # crea data/sample/ con 6 squadre finte
python -m rosicom                     # apre la pagina nel browser
```

## Usalo con la tua lega

### 1. Trova i dati della tua lega

Accedi a Leghe Fantacalcio dal browser e apri la **classifica** della competizione che vuoi seguire
(il campionato, non la coppa). L'indirizzo è di questo tipo:

```
https://leghe.fantacalcio.it/nomedellatualega/classifica?id=123456
```

- **alias** della lega: la parte dopo `leghe.fantacalcio.it/` → `nomedellatualega`
- **id della competizione**: il numero dopo `?id=` → `123456`
  (se non c'è, scegli la competizione dal menu in alto e l'indirizzo si aggiorna)
- **nome della competizione**: come appare sul sito, es. `NomeDellaTuaLega`

### 2. Crea il file `.env`

```bash
cp .env.example .env
```

Apri `.env` e compilalo:

```
LEGHE_USERNAME=il_tuo_username
LEGHE_PASSWORD=la_tua_password
LEGHE_APP_KEY=...                # lascialo com'è
LEGHE_ALIAS=nomedellatualega
LEGHE_ID_COMPETIZIONE=123456
LEGHE_NOME_COMPETIZIONE=NomeDellaTuaLega
```

`.env` contiene la tua password: resta solo sul tuo computer ed è escluso da git (`.gitignore`).
Non condividerlo e non committarlo.

### 3. Scarica i dati

```bash
python scripts/scarica_leghe.py
```

Lo script fa il login, scarica gli export Excel ufficiali di rose e calendario
(gli stessi dei pulsanti "XLSX" e "Scarica ora" sul sito) e li converte in:

```
data/nomedellatualega/
├── raw/rose.xlsx, raw/calendario.xlsx   # file originali
├── squadre.csv                          # squadra, giocatore, ruolo, club, costo
└── partite.csv                          # giornata, casa, trasferta, punti, gol
```

### 4. Apri la pagina

```bash
python -m rosicom
```

Si apre il browser con squadre, classifica e calendario. Se `data/<alias>/` esiste
viene usata la tua lega, altrimenti i dati di esempio. La pagina resta salvata in `output/report.html`.

Ogni giornata: rilancia i passi 3 e 4.

### Penalità

L'export di Leghe non contiene le penalità. Se nella tua lega ce ne sono, scrivile a mano in
`data/<alias>/penalita.csv`:

```
squadra,punti
NomeSquadra,-3
```

Controlla la classifica con quella del sito: se una squadra ha punti diversi, di solito è una penalità mancante.

## Comandi utili

```bash
python -m rosicom data/altralega          # usa un'altra cartella di dati
python -m rosicom --out output/prova      # salva la pagina altrove
python -m rosicom --no-browser            # solo testo nel terminale
python -m pytest                          # test automatici
```

## Inserire i dati a mano

Senza Leghe puoi scrivere i CSV da solo in una cartella e lanciare `python -m rosicom <cartella>`:

- `squadre.csv`: `squadra,giocatore,ruolo` (ruolo P/D/C/A)
- `partite.csv`: `giornata,casa,trasferta,punti_casa,punti_trasferta[,gol_casa,gol_trasferta]`
  — lascia vuoti i punti delle partite da giocare. Senza colonne dei gol, i gol si calcolano dai fantapunti
  con le fasce in `rosicom/config.py` (66 = 1 gol, poi +1 ogni 6 punti).

## Limiti

- Leghe Fantacalcio non ha un'API pubblica: lo script usa gli stessi indirizzi del sito.
  Se il sito cambia, il download può smettere di funzionare.
- Le condizioni d'uso di Fantacalcio potrebbero non permettere lo scaricamento automatico:
  usalo per uso personale e con moderazione.
- Viene seguita una competizione per volta (di solito il campionato); le coppe a gironi non sono supportate.

## Struttura

```
rosicom/
├── rosicom/            # codice: lettura dati, classifica, pagina HTML, conversione Excel di Leghe
├── scripts/            # scarica_leghe.py, genera_esempio.py
├── tests/              # test pytest
├── data/sample/        # lega di esempio
└── output/             # pagina generata (non versionata)
```
