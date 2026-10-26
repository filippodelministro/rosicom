import pytest

from rosicom import loader


def test_csv_con_punto_e_virgola_e_decimali_italiani(tmp_path):
    (tmp_path / "partite.csv").write_text(
        "giornata;casa;trasferta;punti_casa;punti_trasferta\n1;A;B;72,5;65\n")
    df = loader.carica_partite(tmp_path)
    assert df.loc[0, "punti_casa"] == 72.5


def test_colonne_mancanti(tmp_path):
    (tmp_path / "partite.csv").write_text("giornata,casa\n1,A\n")
    with pytest.raises(ValueError, match="colonne mancanti"):
        loader.carica_partite(tmp_path)


def test_file_mancante(tmp_path):
    with pytest.raises(SystemExit, match="File non trovato"):
        loader.carica_partite(tmp_path)


def test_formazioni_opzionali(tmp_path):
    assert loader.carica_formazioni(tmp_path) is None


def test_ruolo_sconosciuto(tmp_path):
    (tmp_path / "formazioni.csv").write_text(
        "giornata,squadra,giocatore,ruolo,voto,fantavoto\n1,A,x,Z,6,6\n")
    with pytest.raises(ValueError, match="ruoli sconosciuti"):
        loader.carica_formazioni(tmp_path)


def test_partite_per_squadra():
    import pandas as pd
    df = pd.DataFrame([(1, "A", "B", 70, 60)], columns=loader.COLONNE_PARTITE)
    r = loader.partite_per_squadra(df).set_index("squadra")
    assert r.loc["B", "fatti"] == 60 and r.loc["B", "subiti"] == 70
