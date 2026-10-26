"""Scarica rose e calendario da Leghe Fantacalcio e li converte nei CSV di rosicom.

Uso: python scripts/scarica_leghe.py
Legge le credenziali da .env (vedi .env.example). Il file .env non va mai su GitHub.
"""
import base64
import http.cookiejar
import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

RADICE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RADICE))
from rosicom import leghe  # noqa: E402

BASE = "https://leghe.fantacalcio.it/"


def leggi_env() -> dict[str, str]:
    path = RADICE / ".env"
    if not path.exists():
        sys.exit("Manca il file .env: copia .env.example in .env e compilalo.")
    env = {}
    for riga in path.read_text().splitlines():
        if "=" in riga and not riga.lstrip().startswith("#"):
            k, v = riga.split("=", 1)
            env[k.strip()] = v.strip()
    mancanti = [k for k in ("LEGHE_USERNAME", "LEGHE_PASSWORD", "LEGHE_APP_KEY", "LEGHE_ALIAS",
                            "LEGHE_ID_COMPETIZIONE") if not env.get(k)]
    if mancanti:
        sys.exit(f".env: valori mancanti {mancanti}")
    return env


def main() -> None:
    env = leggi_env()
    alias = env["LEGHE_ALIAS"]
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    intestazioni = {"app_key": env["LEGHE_APP_KEY"], "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0"}

    corpo = json.dumps({"username": env["LEGHE_USERNAME"], "password": env["LEGHE_PASSWORD"]}).encode()
    req = urllib.request.Request(f"{BASE}api/v1/v1_utente/login?alias_lega={alias}",
                                 data=corpo, headers=intestazioni, method="PUT")
    with opener.open(req) as r:
        risposta = json.loads(r.read())
    if "data" in risposta:  # il sito codifica la risposta del login in base64
        risposta = json.loads(base64.b64decode(risposta["data"]))
    if not risposta.get("success"):
        sys.exit(f"Login fallito: {risposta.get('error_msgs') or risposta}")
    print("Login riuscito")

    cartella = RADICE / "data" / alias
    raw = cartella / "raw"
    raw.mkdir(parents=True, exist_ok=True)
    file = {
        "rose.xlsx": f"servizi/v1_legheSquadra/excel?alias_lega={alias}",
        "calendario.xlsx": "servizi/v1_legheCompetizione/excel?" + urllib.parse.urlencode(
            {"alias_lega": alias, "id_competizione": env["LEGHE_ID_COMPETIZIONE"],
             "nome_competizione": env.get("LEGHE_NOME_COMPETIZIONE", alias)}),
    }
    for nome, url in file.items():
        with opener.open(urllib.request.Request(BASE + url, headers=intestazioni)) as r:
            (raw / nome).write_bytes(r.read())
        print(f"Scaricato {raw / nome}")

    leghe.scrivi_csv(leghe.rose(leghe.griglia_da_xlsx(raw / "rose.xlsx")), cartella / "squadre.csv")
    leghe.scrivi_csv(leghe.calendario(leghe.griglia_da_xlsx(raw / "calendario.xlsx")), cartella / "partite.csv")
    print(f"CSV aggiornati in {cartella}/  ->  python -m rosicom {cartella.relative_to(RADICE)}")


if __name__ == "__main__":
    main()
