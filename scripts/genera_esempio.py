"""Genera una lega finta in data/sample/: rose, calendario completo e risultati delle prime giornate."""
import csv
import random
from pathlib import Path

random.seed(7)
SQUADRE = ["Rosiconi", "Real Sfiga", "Atletico Ma Non Troppo", "Dinamo Divano",
           "Sporting Panchina", "Inter Nos"]
ROSA = {"P": 3, "D": 8, "C": 8, "A": 6}
NOMI = ["Rossi", "Bianchi", "Esposito", "Romano", "Colombo", "Ricci", "Marino", "Greco",
        "Bruno", "Gallo", "Conti", "De Luca", "Mancini", "Costa", "Giordano", "Rizzo",
        "Lombardi", "Moretti", "Barbieri", "Fontana", "Santoro", "Mariani", "Rinaldi",
        "Caruso", "Ferrara", "Galli", "Martini", "Leone", "Longo", "Gentile", "Martinelli",
        "Vitale", "Lombardo", "Serra", "Coppola", "De Santis", "D'Angelo", "Marchetti",
        "Parisi", "Villa", "Conte", "Ferraro", "Ferri", "Fabbri", "Bianco", "Marini",
        "Grasso", "Valentini", "Messina", "Sala", "De Angelis", "Gatti", "Pellegrini",
        "Palumbo", "Sanna", "Farina", "Rizzi", "Monti", "Cattaneo", "Morelli", "Amato",
        "Silvestri", "Mazza", "Testa", "Grassi", "Pellegrino", "Carbone", "Giuliani",
        "Benedetti", "Barone", "Rossetti", "Caputo", "Montanari", "Guerra", "Palmieri",
        "Bernardi", "Martino", "Fiore", "De Rosa", "Ferretti", "Bellini", "Basile",
        "Riva", "Donati", "Piras", "Vitali", "Battaglia", "Sartori", "Neri", "Costantini",
        "Milani", "Pagano", "Ruggiero", "Sorrentino", "D'Amico", "Orlando", "Damico",
        "Negri", "Mele", "Bassi", "Fumagalli", "Pace", "Cirillo", "Bonetti", "Marras",
        "Russo", "Gentili", "Cocco", "Pozzi", "Ruggeri", "Cristiani", "Ventura", "Bertolini",
        "Bonelli", "Caputi", "Melis", "Carli", "Lodi", "Bosco", "Esposti", "Tosi", "Rota",
        "Zanetti", "Gallo", "Farinelli", "Volpe", "Cerri", "Sorrenti", "Lanza", "Bassani",
        "Ricciardi", "Pastore", "Alberti", "Benini", "Fioravanti", "Masi", "Nardi", "Merlo",
        "Pace", "Tedesco", "Olivieri", "Fadda", "Pisani", "Brambilla", "Cortese", "Ferrero",
        "Moro", "Righi", "Sacco", "Locatelli", "Belotti", "Poli", "Ghezzi", "Rocca",
        "Salvi", "Corsi", "Bassetti", "Mosca", "Valente", "Carta", "Spinelli", "Di Marco"]
GIORNATE_GIOCATE = 6
out = Path(__file__).resolve().parent.parent / "data" / "sample"
out.mkdir(parents=True, exist_ok=True)

nomi = random.sample(sorted(set(NOMI)), len(SQUADRE) * sum(ROSA.values()))
rose = []
for s in SQUADRE:
    for ruolo, n in ROSA.items():
        rose += [[s, nomi.pop(), ruolo] for _ in range(n)]

# calendario all'italiana (metodo del cerchio), andata e ritorno
squadre = SQUADRE[:]
random.shuffle(squadre)
andata = []
for _ in range(len(squadre) - 1):
    meta = len(squadre) // 2
    andata.append(list(zip(squadre[:meta], reversed(squadre[meta:]))))
    squadre = [squadre[0], squadre[-1], *squadre[1:-1]]
calendario = andata + [[(b, a) for a, b in g] for g in andata]

forza = {s: random.uniform(-4, 4) for s in SQUADRE}
partite = []
for g, accoppiamenti in enumerate(calendario, start=1):
    for casa, trasf in accoppiamenti:
        if g <= GIORNATE_GIOCATE:
            pc = round(random.gauss(72 + forza[casa], 5) * 2) / 2
            pt = round(random.gauss(70 + forza[trasf], 5) * 2) / 2
        else:
            pc = pt = ""
        partite.append([g, casa, trasf, pc, pt])

with open(out / "squadre.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["squadra", "giocatore", "ruolo"])
    w.writerows(rose)
with open(out / "partite.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["giornata", "casa", "trasferta", "punti_casa", "punti_trasferta"])
    w.writerows(partite)
(out / "formazioni.csv").unlink(missing_ok=True)
print(f"Lega di esempio scritta in {out}")
