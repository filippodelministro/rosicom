"""Pagina HTML locale con squadre, classifica e calendario."""
from html import escape
from pathlib import Path

import pandas as pd

from .config import RUOLI

NOMI_RUOLI = {"P": "Portieri", "D": "Difensori", "C": "Centrocampisti", "A": "Attaccanti"}

STILE = """
:root{--bg:#f6f7f9;--card:#fff;--txt:#1c1e21;--muted:#666;--alt:#f2f4f7;--head:#1f3b57;--line:#e3e6ea}
@media (prefers-color-scheme:dark){:root{--bg:#16181c;--card:#22252b;--txt:#e6e6e6;--muted:#999;
 --alt:#2a2e35;--head:#2c4f73;--line:#33373e}}
body{font-family:-apple-system,system-ui,sans-serif;background:var(--bg);color:var(--txt);
     max-width:960px;margin:0 auto;padding:24px 16px}
h1{margin:0}h2{margin:36px 0 12px}p.nota{color:var(--muted);margin:4px 0 0}
nav a{margin-right:16px;color:inherit}
.card{background:var(--card);border-radius:10px;box-shadow:0 1px 3px #0002;overflow-x:auto}
table{border-collapse:collapse;width:100%}
th,td{padding:8px 12px;text-align:right;white-space:nowrap}
th.l,td.l{text-align:left}thead th{background:var(--head);color:#fff}
tbody tr:nth-child(even){background:var(--alt)}
.griglia{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px}
.griglia .card{padding:12px 16px}.griglia h3{margin:0 0 8px}
.ruolo{color:var(--muted);font-size:13px;margin:8px 0 2px}.ruolo+div{font-size:14px;line-height:1.5}
.giornata{margin-bottom:12px}.giornata h3{margin:0;padding:8px 12px;font-size:15px;border-bottom:1px solid var(--line)}
.giornata td.ris{text-align:center;font-weight:600;width:70px}.giornata td.fp{color:var(--muted);font-size:13px;width:130px}
.da{color:var(--muted)}
"""


def _squadre(rose: pd.DataFrame | None) -> str:
    if rose is None:
        return "<p class='nota'>Nessun file squadre.csv.</p>"
    schede = []
    for squadra, rosa in rose.groupby("squadra", sort=True):
        reparti = "".join(
            f"<div class='ruolo'>{NOMI_RUOLI[r]}</div><div>"
            + ", ".join(escape(g) for g in rosa[rosa["ruolo"] == r]["giocatore"]) + "</div>"
            for r in RUOLI if (rosa["ruolo"] == r).any()
        )
        schede.append(f"<div class='card'><h3>{escape(squadra)}</h3>{reparti}</div>")
    return f"<div class='griglia'>{''.join(schede)}</div>"


def _classifica(tab: pd.DataFrame) -> str:
    colonne = [("pt", "Pt"), ("g", "G"), ("v", "V"), ("n", "N"), ("p", "P"),
               ("gf", "GF"), ("gs", "GS"), ("fp_fatti", "Fantapunti")]
    testa = "<th>#</th><th class='l'>Squadra</th>" + "".join(f"<th>{t}</th>" for _, t in colonne)
    righe = "".join(
        f"<tr><td>{i}</td><td class='l'><b>{escape(sq)}</b>"
        + (f" <span class='da'>({r['penalita']:g})</span>" if r["penalita"] else "") + "</td>"
        + "".join(f"<td>{r[c]:g}</td>" for c, _ in colonne) + "</tr>"
        for i, (sq, r) in enumerate(tab.iterrows(), start=1)
    )
    return f"<div class='card'><table><thead><tr>{testa}</tr></thead><tbody>{righe}</tbody></table></div>"


def _calendario(partite: pd.DataFrame) -> str:
    blocchi = []
    for g, partite_g in partite.groupby("giornata"):
        giocata = partite_g[["punti_casa", "punti_trasferta"]].notna().all(axis=1).all()
        stato = "" if giocata else " <span class='da'>· da giocare</span>"
        righe = []
        for _, p in partite_g.iterrows():
            if pd.notna(p["punti_casa"]):
                ris = f"{p['gol_casa']:g} – {p['gol_trasferta']:g}"
                fp = f"{p['punti_casa']:g} – {p['punti_trasferta']:g}"
            else:
                ris, fp = "<span class='da'>–</span>", ""
            righe.append(f"<tr><td>{escape(p['casa'])}</td><td class='ris'>{ris}</td>"
                         f"<td class='l'>{escape(p['trasferta'])}</td><td class='fp'>{fp}</td></tr>")
        blocchi.append(f"<div class='card giornata'><h3>Giornata {g}{stato}</h3>"
                       f"<table><tbody>{''.join(righe)}</tbody></table></div>")
    return "".join(blocchi)


def salva_html(rose, classifica, partite, cartella: Path, titolo: str) -> Path:
    cartella.mkdir(parents=True, exist_ok=True)
    corpo = (f"<h1>Rosicom</h1><p class='nota'>Dati: {escape(titolo)}</p>"
             "<nav><a href='#squadre'>Squadre</a><a href='#classifica'>Classifica</a>"
             "<a href='#calendario'>Calendario</a></nav>"
             f"<h2 id='squadre'>Squadre</h2>{_squadre(rose)}"
             f"<h2 id='classifica'>Classifica</h2>{_classifica(classifica)}"
             f"<h2 id='calendario'>Calendario</h2>{_calendario(partite)}")

    path = cartella / "report.html"
    path.write_text(f"<!doctype html><html lang='it'><head><meta charset='utf-8'>"
                    f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
                    f"<title>Rosicom</title><style>{STILE}</style></head><body>{corpo}</body></html>")
    return path
