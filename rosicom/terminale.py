"""Interfaccia a terminale: classifica, calendario e squadre stampate con rich."""
import pandas as pd
from rich import box
from rich.columns import Columns
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from .config import RUOLI

NOMI_RUOLI = {"P": "Portieri", "D": "Difensori", "C": "Centrocampisti", "A": "Attaccanti"}
console = Console()


def _n(x) -> str:
    return f"{x:g}"


def classifica(tab: pd.DataFrame) -> None:
    t = Table(title="Classifica", box=box.SIMPLE_HEAVY, title_style="bold")
    t.add_column("#", justify="right", style="dim")
    t.add_column("Squadra", style="bold")
    for nome in ("Pt", "G", "V", "N", "P", "GF", "GS", "Fantapunti"):
        t.add_column(nome, justify="right", style="bold cyan" if nome == "Pt" else None)
    for i, (sq, r) in enumerate(tab.iterrows(), start=1):
        nome = sq + (f" [red]({_n(r['penalita'])})[/red]" if r["penalita"] else "")
        t.add_row(str(i), nome, *(_n(r[c]) for c in ("pt", "g", "v", "n", "p", "gf", "gs", "fp_fatti")))
    console.print(t)


def _giocata(partite_g: pd.DataFrame) -> bool:
    return bool(partite_g["gol_casa"].notna().all())


def prossima_giornata(partite: pd.DataFrame) -> int:
    """Prima giornata non ancora giocata (o l'ultima, a campionato finito)."""
    da_giocare = partite[partite["gol_casa"].isna()]["giornata"]
    return int(da_giocare.min()) if len(da_giocare) else int(partite["giornata"].max())


def giornata(partite: pd.DataFrame, n: int) -> None:
    partite_g = partite[partite["giornata"] == n]
    if partite_g.empty:
        console.print(f"[red]La giornata {n} non esiste.[/red]")
        return
    stato = "" if _giocata(partite_g) else " [dim]· da giocare[/dim]"
    t = Table(title=f"Giornata {n}{stato}", box=box.SIMPLE, show_header=False, title_style="bold")
    t.add_column(justify="right")
    t.add_column(justify="center", style="bold")
    t.add_column()
    t.add_column(justify="right", style="dim")
    for _, p in partite_g.iterrows():
        if pd.notna(p["gol_casa"]):
            ris = f"{_n(p['gol_casa'])} - {_n(p['gol_trasferta'])}"
            fp = f"{_n(p['punti_casa'])} - {_n(p['punti_trasferta'])}"
        else:
            ris, fp = "[dim]-[/dim]", ""
        t.add_row(p["casa"], ris, p["trasferta"], fp)
    console.print(t)


def calendario(partite: pd.DataFrame) -> None:
    """Riepilogo: una riga per giornata, con quante partite sono state giocate."""
    t = Table(title="Calendario", box=box.SIMPLE_HEAVY, title_style="bold")
    t.add_column("Giornata", justify="right")
    t.add_column("Stato")
    prossima = prossima_giornata(partite)
    for n, partite_g in partite.groupby("giornata"):
        if _giocata(partite_g):
            stato = "[green]giocata[/green]"
        elif n == prossima:
            stato = "[bold yellow]prossima[/bold yellow]"
        else:
            stato = "[dim]da giocare[/dim]"
        t.add_row(str(n), stato)
    console.print(t)


def squadre(rose: pd.DataFrame | None, nome: str | None = None) -> None:
    if rose is None:
        console.print("[red]Nessun file squadre.csv.[/red]")
        return
    elenco = sorted(rose["squadra"].unique(), key=str.lower)
    if nome:
        trovate = [s for s in elenco if nome.lower() in s.lower()]
        if not trovate:
            console.print(f"[red]Nessuna squadra contiene '{nome}'.[/red]")
            return
        elenco = trovate
    pannelli = []
    for sq in elenco:
        rosa = rose[rose["squadra"] == sq]
        righe = []
        for r in RUOLI:
            giocatori = rosa[rosa["ruolo"] == r]["giocatore"].tolist()
            if giocatori:
                righe.append(f"[dim]{NOMI_RUOLI[r]}[/dim]\n" + ", ".join(giocatori))
        pannelli.append(Panel("\n".join(righe), title=f"[bold]{sq}[/bold]", width=48))
    console.print(Columns(pannelli))
