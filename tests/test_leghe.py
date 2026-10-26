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
