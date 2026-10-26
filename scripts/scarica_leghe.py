"""Scarica rose e calendario della tua lega da Leghe Fantacalcio e apre la pagina di rosicom.

Uso: python scripts/scarica_leghe.py

La prima volta chiede nome della lega, username e password, trova da solo il campionato
e salva le impostazioni in .env (che resta solo sul tuo computer). Le volte dopo basta rilanciarlo.
"""
import base64
import getpass
import http.cookiejar
import json
import ssl
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import certifi

RADICE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RADICE))
from rosicom import leghe  # noqa: E402

BASE = "https://leghe.fantacalcio.it/"
APP_KEY = "ICiELOObd5DF5uJEATi77CRvHiiRuMU0"  # chiave pubblica del sito, uguale per tutti
ENV = RADICE / ".env"


def leggi_env() -> dict[str, str]:
    env = {}
    if ENV.exists():
        for riga in ENV.read_text().splitlines():
            if "=" in riga and not riga.lstrip().startswith("#"):
                k, v = riga.split("=", 1)
                env[k.strip()] = v.strip()
    return {k: v for k, v in env.items() if v}


def scrivi_env(env: dict[str, str]) -> None:
    righe = ["# Impostazioni di rosicom: resta solo sul tuo computer, non condividerlo."]
    righe += [f"{k}={v}" for k, v in env.items()]
    ENV.write_text("\n".join(righe) + "\n")
    ENV.chmod(0o600)


def chiedi(domanda: str) -> str:
    risposta = ""
    while not risposta:
        risposta = input(domanda).strip()
    return risposta


class Leghe:
    def __init__(self, app_key: str):
        jar = urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
        # certificati di certifi: il Python di python.org su Mac spesso non trova quelli di sistema
        https = urllib.request.HTTPSHandler(context=ssl.create_default_context(cafile=certifi.where()))
        self.opener = urllib.request.build_opener(jar, https)
        self.intestazioni = {"app_key": app_key, "Content-Type": "application/json",
                             "User-Agent": "Mozilla/5.0"}

    def _apri(self, url: str, **kw) -> bytes:
        req = urllib.request.Request(BASE + url, headers=self.intestazioni, **kw)
        with self.opener.open(req) as r:
            return r.read()

    def login(self, alias: str, username: str, password: str) -> None:
        corpo = json.dumps({"username": username, "password": password}).encode()
        try:
            risposta = json.loads(self._apri(f"api/v1/v1_utente/login?alias_lega={alias}",
                                             data=corpo, method="PUT"))
        except urllib.error.HTTPError as e:
            sys.exit(f"Login fallito (errore {e.code}). Controlla username e password.")
        if "data" in risposta:  # il sito codifica la risposta del login in base64
            risposta = json.loads(base64.b64decode(risposta["data"]))
        if not risposta.get("success"):
            sys.exit(f"Login fallito: {risposta.get('error_msgs') or 'username o password errati'}")

    def competizioni(self, alias: str) -> list[dict]:
        try:
            html = self._apri(f"{alias}/classifica").decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            sys.exit(f"Lega '{alias}' non trovata (errore {e.code}). Controlla il nome.")
        return leghe.competizioni(html)

    def scarica(self, url: str, path: Path) -> None:
        path.write_bytes(self._apri(url))


def scegli_campionato(elenco: list[dict]) -> dict:
    if not elenco:
        sys.exit("Non trovo competizioni nella lega: sei sicuro di farne parte?")
    campionati = [c for c in elenco if c["tipo"] == leghe.TIPO_CAMPIONATO]
    if len(campionati) == 1:
        return campionati[0]
    print("Quale competizione vuoi seguire?")
    for i, c in enumerate(elenco, start=1):
        print(f"  {i}. {c['nome']}")
    while True:
        scelta = chiedi("Numero: ")
        if scelta.isdigit() and 1 <= int(scelta) <= len(elenco):
            return elenco[int(scelta) - 1]


def main() -> None:
    env = leggi_env()
    prima = dict(env)
    prima_volta = "LEGHE_ALIAS" not in env
    if prima_volta:
        print("Configurazione di rosicom (solo la prima volta)\n")
        env["LEGHE_ALIAS"] = leghe.alias_da_testo(chiedi("Nome della lega (o link della lega): "))
    alias = env["LEGHE_ALIAS"]
    if "LEGHE_USERNAME" not in env:
        env["LEGHE_USERNAME"] = chiedi("Username Leghe Fantacalcio: ")
    password = env.get("LEGHE_PASSWORD") or getpass.getpass("Password (non viene mostrata): ")

    sito = Leghe(env.get("LEGHE_APP_KEY", APP_KEY))
    sito.login(alias, env["LEGHE_USERNAME"], password)
    print("Login riuscito")

    if "LEGHE_ID_COMPETIZIONE" not in env:
        comp = scegli_campionato(sito.competizioni(alias))
        env["LEGHE_ID_COMPETIZIONE"], env["LEGHE_NOME_COMPETIZIONE"] = comp["id"], comp["nome"]
        print(f"Competizione: {comp['nome']}")
    if "LEGHE_PASSWORD" not in env and (prima_volta or env != prima):
        if input("Salvare la password in .env per non doverla riscrivere? [s/N] ").strip().lower() == "s":
            env["LEGHE_PASSWORD"] = password
    if env != prima:
        scrivi_env(env)
        print(f"Impostazioni salvate in {ENV.name}")

    cartella = RADICE / "data" / alias
    raw = cartella / "raw"
    raw.mkdir(parents=True, exist_ok=True)
    sito.scarica(f"servizi/v1_legheSquadra/excel?alias_lega={alias}", raw / "rose.xlsx")
    sito.scarica("servizi/v1_legheCompetizione/excel?" + urllib.parse.urlencode(
        {"alias_lega": alias, "id_competizione": env["LEGHE_ID_COMPETIZIONE"],
         "nome_competizione": env["LEGHE_NOME_COMPETIZIONE"]}), raw / "calendario.xlsx")
    rose = leghe.rose(leghe.griglia_da_xlsx(raw / "rose.xlsx"))
    partite = leghe.calendario(leghe.griglia_da_xlsx(raw / "calendario.xlsx"))
    if not rose or not partite:
        sys.exit("I file scaricati non hanno il formato atteso: forse il sito è cambiato.")
    leghe.scrivi_csv(rose, cartella / "squadre.csv")
    leghe.scrivi_csv(partite, cartella / "partite.csv")
    print(f"Scaricate {len({g['squadra'] for g in rose})} squadre e {len(partite)} partite\n")

    subprocess.run([sys.executable, "-m", "rosicom", "--dati", str(cartella)], cwd=RADICE)


if __name__ == "__main__":
    main()
