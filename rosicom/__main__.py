"""Uso: python -m rosicom [cartella_dati] [--out cartella_output] [--no-browser]"""
import argparse
import webbrowser
from pathlib import Path

from . import loader, metrics, report


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


def main() -> None:
    ap = argparse.ArgumentParser(prog="rosicom")
    ap.add_argument("dati", nargs="?", default=cartella_predefinita(), type=Path)
    ap.add_argument("--out", default="output", type=Path)
    ap.add_argument("--no-browser", action="store_true", help="non aprire la pagina nel browser")
    args = ap.parse_args()

    partite = loader.carica_partite(args.dati)
    classifica = metrics.classifica(loader.partite_per_squadra(loader.giocate(partite)),
                                    loader.carica_penalita(args.dati))
    rose = loader.carica_squadre(args.dati)

    print(classifica[["pt", "g", "v", "n", "p", "gf", "gs"]].to_string())
    pagina = report.salva_html(rose, classifica, partite, args.out, str(args.dati))
    print(f"\nPagina salvata in {pagina}")
    if not args.no_browser:
        webbrowser.open(pagina.resolve().as_uri())


if __name__ == "__main__":
    main()
