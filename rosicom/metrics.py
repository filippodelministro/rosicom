"""Metriche: classifica reale, classifica 'tutti contro tutti', indice di sfortuna, reparti."""
import math

import pandas as pd

from . import config as cfg


def gol(punti: float) -> int:
    if punti < cfg.SOGLIA_PRIMO_GOL:
        return 0
    return 1 + math.floor((punti - cfg.SOGLIA_PRIMO_GOL) / cfg.AMPIEZZA_FASCIA)


def esito(fatti: float, subiti: float) -> int:
    return esito_gol(gol(fatti), gol(subiti))


def esito_gol(gf: int, gs: int) -> int:
    if gf > gs:
        return cfg.PUNTI_VITTORIA
    if gf == gs:
        return cfg.PUNTI_PAREGGIO
    return cfg.PUNTI_SCONFITTA


def classifica(righe: pd.DataFrame, penalita: dict[str, float] | None = None) -> pd.DataFrame:
    r = righe.copy()
    if "gf" not in r:  # senza gol ufficiali li calcolo dalle fasce di config.py
        r["gf"] = r["fatti"].map(gol)
        r["gs"] = r["subiti"].map(gol)
    r["pt"] = [esito_gol(f, s) for f, s in zip(r["gf"], r["gs"])]
    tab = r.groupby("squadra").agg(
        g=("pt", "size"),
        v=("pt", lambda x: (x == cfg.PUNTI_VITTORIA).sum()),
        n=("pt", lambda x: (x == cfg.PUNTI_PAREGGIO).sum()),
        p=("pt", lambda x: (x == cfg.PUNTI_SCONFITTA).sum()),
        gf=("gf", "sum"), gs=("gs", "sum"),
        pt=("pt", "sum"),
        fp_fatti=("fatti", "sum"), fp_subiti=("subiti", "sum"),
    )
    tab["media_fp"] = (tab["fp_fatti"] / tab["g"]).round(2)
    tab["penalita"] = tab.index.map(penalita or {}).fillna(0)
    tab["pt"] = tab["pt"] + tab["penalita"]
    return tab.sort_values(["pt", "fp_fatti"], ascending=False)


def punti_attesi(righe: pd.DataFrame) -> pd.Series:
    """Per ogni giornata, punti medi che la squadra avrebbe fatto contro TUTTE le altre."""
    attesi = {}
    for _, g in righe.groupby("giornata"):
        punteggi = dict(zip(g["squadra"], g["fatti"]))
        for sq, fp in punteggi.items():
            altri = [esito(fp, x) for s, x in punteggi.items() if s != sq]
            attesi[sq] = attesi.get(sq, 0) + sum(altri) / len(altri)
    return pd.Series(attesi, name="xpt")


def rosicate(righe: pd.DataFrame) -> pd.Series:
    """Sconfitte con un punteggio sopra la mediana di giornata."""
    r = righe.copy()
    r["mediana"] = r.groupby("giornata")["fatti"].transform("median")
    r["pt"] = [esito(f, s) for f, s in zip(r["fatti"], r["subiti"])]
    mask = (r["pt"] == cfg.PUNTI_SCONFITTA) & (r["fatti"] > r["mediana"])
    return r[mask].groupby("squadra").size().rename("rosicate")


def indice_sfortuna(righe: pd.DataFrame) -> pd.DataFrame:
    tab = classifica(righe)[["g", "pt", "fp_subiti"]]
    tab = tab.join(punti_attesi(righe)).join(rosicate(righe)).fillna({"rosicate": 0})
    tab["xpt"] = tab["xpt"].round(2)
    tab["sfortuna"] = (tab["xpt"] - tab["pt"]).round(2)  # >0 = meriteresti di più
    media_subiti = tab["fp_subiti"].sum() / tab["g"].sum()
    tab["subiti_vs_media"] = (tab["fp_subiti"] / tab["g"] - media_subiti).round(2)
    tab["rosicate"] = tab["rosicate"].astype(int)
    return tab.drop(columns="fp_subiti").sort_values("sfortuna", ascending=False)


def reparti(formazioni: pd.DataFrame) -> pd.DataFrame:
    """Fantavoto medio per giocatore schierato, per squadra e reparto."""
    tab = formazioni.pivot_table(index="squadra", columns="ruolo",
                                 values="fantavoto", aggfunc="mean")
    return tab.reindex(columns=cfg.RUOLI).round(2)
