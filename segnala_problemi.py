#!/usr/bin/env python3
"""Avviso per e-mail quando l'aggiornamento automatico dei dati NON riesce (usato dal workflow dopo aggiorna_dati.py).

Come arriva l'e-mail: il programma apre una «segnalazione» (issue) nel repository GitHub con un testo chiaro in
italiano; GitHub la manda per e-mail al proprietario dell'account (che viene anche citato con @nome).
- C'è un problema e non c'è una segnalazione aperta  -> ne apre una nuova (e-mail «⚠ … non riuscito»).
- Il problema si ripete (segnalazione già aperta)     -> aggiunge un commento (e-mail «si è ripresentato»).
- Tutto a posto e c'è una segnalazione aperta         -> commento «✅ risolto» e la chiude (e-mail di chiusura).
- Tutto a posto e nessuna segnalazione aperta         -> niente (nessuna e-mail).
Se ISTAT non ha ancora pubblicato dati nuovi NON è un problema: aggiorna_dati.py finisce senza problemi
e qui non succede niente.
Solo libreria standard di Python. Variabili d'ambiente del workflow: GITHUB_TOKEN, GITHUB_REPOSITORY,
GITHUB_REPOSITORY_OWNER, GITHUB_SERVER_URL, GITHUB_RUN_ID. Con --stampa scrive i testi invece di chiamare GitHub.
"""
import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone

ETICHETTA = "aggiornamento-non-riuscito"
TITOLO = "⚠ Conti economici: l'aggiornamento automatico dei dati del valore aggiunto NON è riuscito"
# testi della segnalazione
FONTE = {
    "manuale": "Se serve subito un dato più recente, nella pagina si può trascinare il file Excel di Tagliacarne (e i CSV di ISTAT) scaricati a mano (punto 2 della pagina).",
    "nessun_dato_nuovo": "Se il Centro Studi Tagliacarne o ISTAT non hanno ancora pubblicato dati nuovi non arriva nessuna e-mail.",
    "periodo_salvato": "Dati sul sito: fino al {}.",
    "fino": "fino al",
}
MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"]

# cosa vuol dire e cosa fare, per tipo di problema (tipi di aggiorna_dati.py)
SPIEGAZIONI = {
    "rete": ("il sito (Centro Studi Tagliacarne o ISTAT) non ha risposto, nemmeno riprovando più volte.",
             "Di solito è un disservizio temporaneo: non serve fare niente, il controllo viene ripetuto da solo il giorno dopo. "
             "Se l'avviso si ripete per più giorni, controllare che i siti https://www.tagliacarne.it e https://esploradati.istat.it funzionino."),
    "risposta": ("il sito ha risposto, ma invece del file dei dati ha mandato altro (per esempio una pagina di manutenzione o un file non valido).",
                 "Di solito è temporaneo: il controllo viene ripetuto da solo. Se l'avviso si ripete per più giorni, il servizio di download "
                 "potrebbe essere cambiato: serve una modifica al programma."),
    "formato": ("il file o la tavola sono cambiati (intestazioni, colonne, nomi dei territori o link diversi da quelli attesi).",
                "Serve una modifica al programma di aggiornamento (aggiorna_dati.py) e forse alla pagina: vedere HANDOVER §29. "
                "Il problema non si risolve da solo. Intanto la pagina continua a usare gli ultimi dati salvati."),
    "anomalia": ("i dati scaricati non sono plausibili o sono diversi da quelli già salvati (anno più vecchio, valori molto cambiati, "
                 "somme che non tornano, variazione sull'anno prima fuori dalla norma).",
                 "Per sicurezza i dati NON sono stati sostituiti. Se Tagliacarne ha davvero rivisto i dati, controllare il suo comunicato stampa e il file "
                 "Excel e, se è tutto giusto, rilanciare l'aggiornamento da GitHub (Actions → «Aggiorna dati valore aggiunto» → Run workflow, "
                 "spuntando «Accetta»). Dettagli nel testo tecnico qui sotto."),
    "prova": ("questa è una PROVA dell'avviso, lanciata a mano: non c'è nessun problema reale.",
              "Nessuna azione: la prossima esecuzione normale chiuderà da sola questa segnalazione."),
    "imprevisto": ("il programma di aggiornamento ha incontrato un errore imprevisto.",
                   "Serve un controllo del programma (vedere i dettagli tecnici qui sotto e HANDOVER §29)."),
}


def data_it(iso=None):
    """«29 settembre 2026 alle 20:03» (ora italiana)."""
    d = datetime.fromisoformat(iso.replace("Z", "+00:00")) if iso else datetime.now(timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        d = d.astimezone(ZoneInfo("Europe/Rome"))
    except Exception:
        pass
    return f"{d.day} {MESI[d.month - 1]} {d.year} alle {d:%H:%M}"


def mese_it(iso):
    return f"{MESI[int(iso[5:7]) - 1]} {iso[:4]}" if iso else "—"


def problemi_da(esito, passo_scarica, passo_salva):
    """Elenco dei problemi: quelli delle tavole (esito.json) più quelli dei passi del workflow."""
    problemi = list(esito.get("problemi", []))
    if passo_scarica not in ("success", "", None):
        problemi.append({"tavola": "—", "nome": "programma di aggiornamento", "tipo": "imprevisto",
                         "dettaglio": f"il passo «Scarica i dati» è finito con esito «{passo_scarica}»"
                                      + (" (tempo massimo superato o esecuzione annullata)" if passo_scarica == "cancelled" else "")})
    if passo_salva == "failure":
        problemi.append({"tavola": "—", "nome": "salvataggio nel repository", "tipo": "imprevisto",
                         "dettaglio": "i dati scaricati non sono stati salvati nel repository (passo «Salva i dati nel repository» non riuscito)"})
    return problemi


def testo_problemi(problemi, esito, sito, link_run, proprietario, ripetuto=False):
    quando = data_it(esito.get("quando"))
    righe = []
    if ripetuto:
        righe.append(f"@{proprietario} **il problema si è ripresentato** nel controllo del {quando}.\n")
    else:
        righe.append(f"@{proprietario} l'aggiornamento automatico dei dati del **sito Conti economici (valore aggiunto)** del {quando} **non è riuscito**.\n")
    righe.append("### Cosa è successo")
    for p in problemi:
        cosa, _ = SPIEGAZIONI.get(p["tipo"], SPIEGAZIONI["imprevisto"])
        salvato = " " + FONTE["periodo_salvato"].format(p["ultimo_anno_salvato"]) if p.get("ultimo_anno_salvato") else ""
        righe.append(f"- **{p['tavola']}** ({p['nome']}): {cosa}{salvato}")
    riuscite = esito.get("tavole", [])
    if riuscite:
        righe.append(f"\nLe altre tavole ({', '.join(t['tavola'] for t in riuscite)}) sono state controllate normalmente.")
    righe.append("\n### Cosa vuol dire")
    if all(p["tipo"] == "riservato" for p in problemi):
        righe.append(f"I dati sono stati **aggiornati normalmente**: mancano solo le tavole indicate sopra. La pagina {sito} funziona.")
    else:
        righe.append(f"I dati già presenti sul sito **non sono stati cancellati né modificati**: la pagina {sito} continua a funzionare "
                     "con i dati dell'ultimo aggiornamento riuscito. " + FONTE["manuale"])
    righe.append("\n### Cosa fare")
    for tipo in dict.fromkeys(p["tipo"] for p in problemi):
        righe.append(f"- {SPIEGAZIONI.get(tipo, SPIEGAZIONI['imprevisto'])[1]}")
    righe.append("\n### Da sapere")
    righe.append("Questo avviso arriva **solo quando qualcosa non funziona**. " + FONTE["nessun_dato_nuovo"] + " "
                 "Quando un aggiornamento successivo riesce, questa segnalazione viene chiusa da sola e arriva un messaggio «risolto».")
    righe.append(f"\n<details><summary>Dettagli tecnici</summary>\n\nEsecuzione: {link_run}\n")
    for p in problemi:
        righe.append(f"- `{p['tavola']}` [{p['tipo']}]: {p['dettaglio']}")
    righe.append("\n</details>")
    return "\n".join(righe)


def testo_risolto(esito, sito):
    tav = ", ".join(f"{t['tavola']} {FONTE['fino']} {t['ultimo_anno']}" for t in esito.get("tavole", []))
    return (f"✅ **Risolto**: l'aggiornamento automatico del {data_it(esito.get('quando'))} è riuscito per tutte le tavole"
            + (f" ({tav})" if tav else "") + f". Il sito {sito} ha i dati aggiornati. La segnalazione viene chiusa.")


class GitHub:
    def __init__(self, token, repo, server="https://api.github.com"):
        self.token, self.repo, self.server = token, repo, server

    def chiama(self, metodo, percorso, dati=None, ok404=False):
        req = urllib.request.Request(f"{self.server}/repos/{self.repo}{percorso}", method=metodo,
                                     data=None if dati is None else json.dumps(dati).encode(),
                                     headers={"Authorization": f"Bearer {self.token}", "Accept": "application/vnd.github+json",
                                              "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "report-conti-economici"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                corpo = r.read()
                return json.loads(corpo) if corpo else None
        except urllib.error.HTTPError as e:
            if ok404 and e.code == 404:
                return None
            raise

    def segnalazioni_aperte(self):
        return self.chiama("GET", f"/issues?state=open&labels={ETICHETTA}&per_page=10") or []

    def crea_etichetta(self):
        if self.chiama("GET", f"/labels/{ETICHETTA}", ok404=True) is None:
            self.chiama("POST", "/labels", {"name": ETICHETTA, "color": "d93f0b",
                                            "description": "Aggiornamento automatico dei dati non riuscito"})

    def apri(self, titolo, testo):
        return self.chiama("POST", "/issues", {"title": titolo, "body": testo, "labels": [ETICHETTA]})

    def commenta(self, numero, testo):
        return self.chiama("POST", f"/issues/{numero}/comments", {"body": testo})

    def chiudi(self, numero):
        return self.chiama("PATCH", f"/issues/{numero}", {"state": "closed", "state_reason": "completed"})


def gestisci(api, esito, passo_scarica, passo_salva, sito, link_run, proprietario):
    """Decide cosa fare; restituisce una frase per il registro del workflow."""
    problemi = problemi_da(esito, passo_scarica, passo_salva)
    aperte = api.segnalazioni_aperte()
    if problemi:
        if aperte:
            n = aperte[0]["number"]
            api.commenta(n, testo_problemi(problemi, esito, sito, link_run, proprietario, ripetuto=True))
            return f"problema ripetuto: commento alla segnalazione #{n}"
        api.crea_etichetta()
        titolo = TITOLO + (" (PROVA)" if all(p["tipo"] == "prova" for p in problemi) else "")
        r = api.apri(titolo, testo_problemi(problemi, esito, sito, link_run, proprietario))
        return f"problema: aperta la segnalazione #{r['number']}"
    for s in aperte:
        api.commenta(s["number"], testo_risolto(esito, sito))
        api.chiudi(s["number"])
    return f"tutto a posto; chiuse {len(aperte)} segnalazioni" if aperte else "tutto a posto, nessuna segnalazione"


class Stampa:
    """Al posto di GitHub: stampa i testi (per provare in locale)."""
    def __init__(self, aperte=()):
        self.aperte, self.azioni = list(aperte), []

    def segnalazioni_aperte(self):
        return self.aperte

    def crea_etichetta(self):
        self.azioni.append(("etichetta",))

    def apri(self, titolo, testo):
        self.azioni.append(("apri", titolo, testo))
        return {"number": 1}

    def commenta(self, numero, testo):
        self.azioni.append(("commenta", numero, testo))

    def chiudi(self, numero):
        self.azioni.append(("chiudi", numero))


def main():
    arg = argparse.ArgumentParser(description="Segnala per e-mail (issue di GitHub) i problemi dell'aggiornamento")
    arg.add_argument("--esito", required=True)
    arg.add_argument("--scarica", default="success", help="esito del passo di download (success, failure, cancelled…)")
    arg.add_argument("--salva", default="success", help="esito del passo di salvataggio")
    arg.add_argument("--stampa", action="store_true", help="non chiama GitHub: stampa i testi")
    a = arg.parse_args()
    try:
        with open(a.esito, encoding="utf-8") as f:
            esito = json.load(f)
    except (OSError, ValueError):
        esito = {}
    repo = os.environ.get("GITHUB_REPOSITORY", "StatisticaVA/report-conti-economici")
    proprietario = os.environ.get("GITHUB_REPOSITORY_OWNER", repo.split("/")[0])
    sito = f"https://{proprietario.lower()}.github.io/{repo.split('/')[1]}/"
    link_run = f"{os.environ.get('GITHUB_SERVER_URL', 'https://github.com')}/{repo}/actions/runs/{os.environ.get('GITHUB_RUN_ID', '?')}"
    if a.stampa:
        api = Stampa()
        print(gestisci(api, esito, a.scarica, a.salva, sito, link_run, proprietario))
        for az in api.azioni:
            print("\n---", az[0], *az[1:], sep="\n")
        return
    api = GitHub(os.environ["GITHUB_TOKEN"], repo)
    print(gestisci(api, esito, a.scarica, a.salva, sito, link_run, proprietario))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # se la segnalazione non si può fare, il workflow fallisce: arriva comunque l'e-mail di GitHub
        print(f"Segnalazione non riuscita: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)
