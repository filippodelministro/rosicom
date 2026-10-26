import pandas as pd
import pytest

from rosicom import loader, metrics


def righe(*partite):
    df = pd.DataFrame(partite, columns=loader.COLONNE_PARTITE)
    return loader.partite_per_squadra(df)


@pytest.mark.parametrize("punti,attesi", [
    (65.5, 0), (66, 1), (71.5, 1), (72, 2), (78, 3), (40, 0),
])
def test_gol(punti, attesi):
    assert metrics.gol(punti) == attesi


def test_esito():
    assert metrics.esito(72, 65) == 3
    assert metrics.esito(66, 71.5) == 1   # stessa fascia = pareggio
    assert metrics.esito(60, 66) == 0


def test_classifica():
    tab = metrics.classifica(righe((1, "A", "B", 72, 65), (1, "C", "D", 66, 66)))
    assert tab.loc["A", "pt"] == 3 and tab.loc["A", "gf"] == 2
    assert tab.loc["B", "pt"] == 0 and tab.loc["B", "gs"] == 2
    assert tab.loc["C", "n"] == 1
    assert list(tab.index)[0] == "A"


def test_punti_attesi_e_sfortuna():
    tab = metrics.indice_sfortuna(righe((1, "A", "B", 72, 65), (1, "C", "D", 66, 66)))
    assert tab.loc["A", "xpt"] == 3
    assert tab.loc["C", "xpt"] == pytest.approx(1.33)
    assert tab.loc["C", "sfortuna"] == pytest.approx(0.33)
    assert tab.loc["B", "sfortuna"] == 0


def test_rosicata():
    # B fa 75 (sopra la mediana) ma perde contro A che fa 80
    r = righe((1, "A", "B", 80, 75), (1, "C", "D", 60, 55))
    assert metrics.rosicate(r).to_dict() == {"B": 1}


def test_reparti():
    f = pd.DataFrame([
        (1, "A", "x", "P", 6, 6), (1, "A", "y", "A", 6, 9), (1, "A", "z", "A", 6, 7),
    ], columns=loader.COLONNE_FORMAZIONI)
    tab = metrics.reparti(f)
    assert tab.loc["A", "A"] == 8
    assert list(tab.columns) == ["P", "D", "C", "A"]


def test_gol_ufficiali_e_penalita():
    r = righe((1, "A", "B", 87.5, 64)).assign(gf=[5, 0], gs=[0, 5])
    tab = metrics.classifica(r, {"A": -3})
    assert tab.loc["A", "gf"] == 5 and tab.loc["A", "pt"] == 0
