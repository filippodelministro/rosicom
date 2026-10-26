"""Uso: python -m rosicom [cartella_dati] [--out cartella_output] [--no-browser]"""
import argparse
import webbrowser
from pathlib import Path

from . import loader, metrics, report


def main() -> None:
    ap = argparse.ArgumentParser(prog="rosicom")
    ap.add_argument("dati", nargs="?", default="data/sample", type=Path)
    ap.add_argument("--out", default="output", type=Path)
    ap.add_argument("--no-browser", action="store_true", help="non aprire la pagina nel browser")
    args = ap.parse_args()

    partite = loader.carica_partite(args.dati)
    classifica = metrics.classifica(loader.partite_per_squadra(loader.giocate(partite)))
    rose = loader.carica_squadre(args.dati)

    print(classifica[["pt", "g", "v", "n", "p", "gf", "gs"]].to_string())
    pagina = report.salva_html(rose, classifica, partite, args.out, str(args.dati))
    print(f"\nPagina salvata in {pagina}")
    if not args.no_browser:
        webbrowser.open(pagina.resolve().as_uri())


if __name__ == "__main__":
    main()
