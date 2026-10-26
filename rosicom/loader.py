"""Caricamento e validazione dei CSV esportati (o compilati a mano).

squadre.csv     squadra,giocatore,ruolo
partite.csv     giornata,casa,trasferta,punti_casa,punti_trasferta[,gol_casa,gol_trasferta]
                (punti vuoti = da giocare; senza colonne gol, i gol si calcolano da config.py)
formazioni.csv  giornata,squadra,giocatore,ruolo,voto,fantavoto   (opzionale)
"""
from pathlib import Path

import pandas as pd

from .config import RUOLI
from .metrics import gol

COLONNE_SQUADRE = ["squadra", "giocatore", "ruolo"]
COLONNE_PARTITE = ["giornata", "casa", "trasferta", "punti_casa", "punti_trasferta"]
COLONNE_FORMAZIONI = ["giornata", "squadra", "giocatore", "ruolo", "voto", "fantavoto"]


def _leggi(path: Path, colonne: list[str]) -> pd.DataFrame:
    if not path.exists():
        raise SystemExit(f"File non trovato: {path} (crealo o controlla la cartella)")
    # sep=None indovina il separatore: gli export italiani usano spesso ";"
    df = pd.read_csv(path, sep=None, engine="python", decimal=",")
    df.columns = [c.strip().lower() for c in df.columns]
    mancanti = set(colonne) - set(df.columns)
    if mancanti:
        raise ValueError(f"{path.name}: colonne mancanti {sorted(mancanti)}")
    return df[colonne]


def carica_partite(cartella: Path) -> pd.DataFrame:
    path = cartella / "partite.csv"
    colonne_gol = ["gol_casa", "gol_trasferta"]
    df = _leggi(path, COLONNE_PARTITE)
    originale = pd.read_csv(path, sep=None, engine="python", decimal=",")
    for c in ("punti_casa", "punti_trasferta"):
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", "."), errors="coerce")
    if set(colonne_gol) <= set(originale.columns):
        for c in colonne_gol:
            df[c] = pd.to_numeric(originale[c], errors="coerce")
    else:  # gol calcolati con le fasce di config.py
        df["gol_casa"] = df["punti_casa"].map(gol, na_action="ignore")
        df["gol_trasferta"] = df["punti_trasferta"].map(gol, na_action="ignore")
    return df


def carica_squadre(cartella: Path) -> pd.DataFrame | None:
    path = cartella / "squadre.csv"
    if not path.exists():
        return None
    df = _leggi(path, COLONNE_SQUADRE)
    df["ruolo"] = df["ruolo"].str.upper().str.strip()
    return df


def carica_penalita(cartella: Path) -> dict[str, float]:
    """penalita.csv (opzionale): squadra,punti   es. VivailPratino,-3"""
    path = cartella / "penalita.csv"
    if not path.exists():
        return {}
    df = _leggi(path, ["squadra", "punti"])
    return df.groupby("squadra")["punti"].sum().to_dict()


def giocate(partite: pd.DataFrame) -> pd.DataFrame:
    return partite.dropna(subset=["punti_casa", "punti_trasferta"])


def carica_formazioni(cartella: Path) -> pd.DataFrame | None:
    path = cartella / "formazioni.csv"
    if not path.exists():
        return None
    df = _leggi(path, COLONNE_FORMAZIONI)
    df["ruolo"] = df["ruolo"].str.upper().str.strip()
    sconosciuti = set(df["ruolo"]) - set(RUOLI)
    if sconosciuti:
        raise ValueError(f"formazioni.csv: ruoli sconosciuti {sorted(sconosciuti)}")
    return df


def partite_per_squadra(partite: pd.DataFrame) -> pd.DataFrame:
    """Una riga per squadra per giornata: fatti, subiti, avversario."""
    casa = partite.rename(columns={"casa": "squadra", "trasferta": "avversario",
                                   "punti_casa": "fatti", "punti_trasferta": "subiti",
                                   "gol_casa": "gf", "gol_trasferta": "gs"})
    trasf = partite.rename(columns={"trasferta": "squadra", "casa": "avversario",
                                    "punti_trasferta": "fatti", "punti_casa": "subiti",
                                    "gol_trasferta": "gf", "gol_casa": "gs"})
    return pd.concat([casa, trasf], ignore_index=True).sort_values(["giornata", "squadra"])
