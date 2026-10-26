from rosicom import leghe

ROSE = [
    ["Rose lega X"], [],
    ["Alfa", "", "", "", "", "Beta"],
    ["Ruolo", "Calciatore", "Squadra", "Costo", "", "Ruolo", "Calciatore", "Squadra", "Costo"],
    ["P", "Rossi", "Int", "3", "", "P", "Bianchi", "Juv", "1"],
    ["A", "Verdi", "Mil", "40", "", "D", "Neri", "Rom", "5"],
    ["Crediti Residui: 5", "", "", "", "", "Crediti Residui: 2"],
]

CALENDARIO = [
    ["1ª Giornata lega", "", "4ª Giornata serie a", "", "", "", "2ª Giornata lega"],
    ["Alfa", "72", "65.5", "Beta", "2-0", "", "Beta", "0", "0", "Alfa", "-"],
]


def test_rose():
    r = leghe.rose(ROSE)
    assert {(g["squadra"], g["giocatore"], g["ruolo"]) for g in r} == {
        ("Alfa", "Rossi", "P"), ("Alfa", "Verdi", "A"), ("Beta", "Bianchi", "P"), ("Beta", "Neri", "D")}


def test_calendario():
    g1, g2 = leghe.calendario(CALENDARIO)
    assert (g1["giornata"], g1["casa"], g1["punti_trasferta"], g1["gol_casa"]) == (1, "Alfa", "65.5", "2")
    assert (g2["giornata"], g2["casa"], g2["punti_casa"], g2["gol_casa"]) == (2, "Beta", "", "")


def test_alias_da_testo():
    assert leghe.alias_da_testo("PancheLudopatico") == "pancheludopatico"
    assert leghe.alias_da_testo("https://leghe.fantacalcio.it/miaLega/classifica?id=1") == "mialega"


def test_competizioni():
    html = ("<li><a href='#' data-isin='true' data-id='662053'><span class='competition-icon "
            "competition-icon-5'></span>La coppa</a></li><li><a href='#' data-isin='true' "
            "data-id='454999'><span class='competition-icon competition-icon-1'></span>Campionato</a>") * 2
    assert leghe.competizioni(html) == [
        {"id": "662053", "tipo": "5", "nome": "La coppa"},
        {"id": "454999", "tipo": "1", "nome": "Campionato"}]
