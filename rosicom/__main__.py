"""Uso: python -m rosicom [classifica | calendario [N] | squadre [nome]] [--dati cartella]

Senza comando apre un menu interattivo.
"""
import argparse
from pathlib import Path

from . import loader, metrics, terminale
from .terminale import console


def cartella_predefinita() -> Path:
    """La lega indicata in .env (LEGHE_ALIAS) se è già stata scaricata, altrimenti i dati di esempio."""
    env = Path(".env")
    if env.exists():
        for riga in env.read_text().splitlines():
            if riga.startswith("LEGHE_ALIAS="):
                lega = Path("data") / riga.split("=", 1)[1].strip()
                if (lega / "partite.csv").exists():
                    return lega
    return Path("data/sample")


def menu(classifica, partite, rose) -> None:
    while True:
        console.print("\n[bold]1[/bold] Classifica   [bold]2[/bold] Giornata   [bold]3[/bold] Calendario"
                      "   [bold]4[/bold] Squadre   [bold]c[/bold] Pulisci   [bold]q[/bold] Esci")
        scelta = console.input("> ").strip().lower()
        if scelta == "1":
            terminale.classifica(classifica)
        elif scelta == "2":
            prossima = terminale.prossima_giornata(partite)
            n = console.input(f"Quale giornata? [dim](Invio = {prossima})[/dim] ").strip()
            terminale.giornata(partite, int(n) if n.isdigit() else prossima)
        elif scelta == "3":
            terminale.calendario(partite)
        elif scelta == "4":
            nome = console.input("Nome squadra [dim](Invio = tutte)[/dim] ").strip()
            terminale.squadre(rose, nome or None)
        elif scelta in ("c", "clear"):
            console.clear()
        elif scelta in ("q", "esci", ""):
            return


def main() -> None:
    ap = argparse.ArgumentParser(prog="rosicom", description="Classifica, calendario e squadre della lega.")
    ap.add_argument("comando", nargs="?", choices=["classifica", "calendario", "squadre"])
    ap.add_argument("valore", nargs="?", help="numero di giornata (calendario) o nome squadra (squadre)")
    ap.add_argument("--dati", default=cartella_predefinita(), type=Path, help="cartella con i CSV della lega")
    args = ap.parse_args()

    partite = loader.carica_partite(args.dati)
    classifica = metrics.classifica(loader.partite_per_squadra(loader.giocate(partite)),
                                    loader.carica_penalita(args.dati))
    rose = loader.carica_squadre(args.dati)
    console.print(f"[dim]Lega: {args.dati}[/dim]")

    if args.comando == "classifica":
        terminale.classifica(classifica)
    elif args.comando == "calendario":
        if args.valore and args.valore.isdigit():
            terminale.giornata(partite, int(args.valore))
        else:
            terminale.calendario(partite)
    elif args.comando == "squadre":
        terminale.squadre(rose, args.valore)
    else:
        terminale.classifica(classifica)
        terminale.giornata(partite, terminale.prossima_giornata(partite))
        menu(classifica, partite, rose)


if __name__ == "__main__":
    main()
