# rosicom

Strumento locale e minimale per analizzare le statistiche di una lega di Fantacalcio.

## Uso

```bash
pip install -r requirements.txt
python scripts/genera_esempio.py      # dati finti in data/sample/
python -m rosicom                     # analizza data/sample/
python -m rosicom data/lega2026 --out output/lega2026
python -m pytest                      # test automatici
```

## Dati in input (CSV, separatore `,` o `;`)

- `squadre.csv`: `squadra,giocatore,ruolo` (ruolo P/D/C/A)
- `partite.csv`: `giornata,casa,trasferta,punti_casa,punti_trasferta` (fantapunti; lascia vuoti i punti delle partite da giocare)

## Cosa mostra

Una pagina nel browser (`output/report.html`) con:

- **Squadre**: rosa di ogni squadra divisa per ruolo
- **Classifica**: punti, vittorie/pareggi/sconfitte, gol (soglia 66, +1 ogni 6 punti: si cambia in `rosicom/config.py`), fantapunti
- **Calendario**: partite giocate con risultato e fantapunti, e partite da giocare
