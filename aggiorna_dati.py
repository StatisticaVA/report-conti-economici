#!/usr/bin/env python3
"""Aggiorna in dati/ i numeri del VALORE AGGIUNTO usati dalla pagina «Conti economici».

Lo usa l'aggiornamento automatico su GitHub (.github/workflows/aggiorna-dati.yml, ogni giorno), ma si può lanciare
anche a mano:  python3 aggiorna_dati.py [--forza] [--accetta]        Solo libreria standard di Python (niente da installare).

REGOLA DEGLI ANNI (decisa dall'utente il 07/10/2026; è quella usata da Osserva), per un anno N:
  - ultimo anno N          -> file del Centro Studi Tagliacarne (quello appena uscito);
  - anno N-1               -> il file Tagliacarne che avevamo già l'anno prima (archivio: dati/tagliacarne_AAAA.csv, mai sovrascritto
                              da un anno più recente);
  - anni N-2 e precedenti  -> ISTAT (conti economici territoriali, dati/istat_valore_aggiunto.csv).

COSA FA, in due parti indipendenti (se una ha un problema l'altra si aggiorna lo stesso):
  1. TAGLIACARNE. Legge la pagina «Statistiche territoriali» del Centro Studi, trova da sola il link del file Excel
     «Valore aggiunto per provincia e branca di attività economica. Anno XXXX (versione gg-mm-aaaa)» (l'indirizzo del
     file contiene una data di caricamento e cambia: per questo non è scritto nel programma), scarica il file (28 KB),
     lo confronta con quello già salvato (impronta SHA-256) e, se è nuovo, lo controlla:
       - fogli, intestazioni e righe attese (le 12 province lombarde, Lombardia, Italia, regioni, ripartizioni);
       - ogni riga: somma delle branche = totale; Lombardia = somma delle 12 province; regioni = ripartizioni = province;
         Italia = province + «extra-regio» (poche centinaia di milioni);
       - variazione rispetto all'anno prima (archivio): oltre ±8% avviso, oltre ±15% blocco; crescita diversa da quella dell'Italia
         di oltre 4 punti avviso, oltre 10 blocco (probabile base rivista);
     poi lo salva come dati/tagliacarne_AAAA.csv. L'anno più vecchio non viene mai toccato da un anno più nuovo.
  2. ISTAT. Scarica il valore aggiunto provinciale (dataflow 93_498_DF_DCCN_PILT_1, tutte le aree) e la popolazione
     al 1° gennaio, ricava Varese, Lombardia (= somma delle 12 province) e Italia (= somma delle 107 province +
     «extra-regio», come fa Osserva), il valore aggiunto pro-capite si calcola nella pagina: dati/istat_valore_aggiunto.csv.
     Riscarica solo se ISTAT dichiara un aggiornamento nuovo (LAST_UPDATE) o se l'ultimo scarico ha più di 28 giorni.
Se qualcosa non va, i file della parte interessata NON vengono toccati e il problema viene scritto nel file di esito
(--esito): segnala_problemi.py apre una segnalazione (e-mail). Se non ci sono novità non è un problema.
  --forza     scarica e ricontrolla tutto anche se non ci sono novità
  --accetta   salva anche se un controllo di plausibilità (variazioni, scarto con il file già salvato) segnala un'anomalia
              (da usare solo dopo aver verificato a mano che il cambiamento è giusto)
Con la variabile d'ambiente PROVA_ERRORE=true si simula un problema (per provare l'e-mail).
"""
import argparse
import csv
import hashlib
import html
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone

CARTELLA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dati")
PAGINA_TAGLIACARNE = "https://www.tagliacarne.it/linee_di_attivita-33/statistiche_territoriali-101/"
SDMXWS = "https://esploradati.istat.it/SDMXWS/rest/"
FLUSSO_VA = "IT1,93_498_DF_DCCN_PILT_1,1.0"
FLUSSO_POP = "IT1,22_289_DF_DCIS_POPRES1_1,1.0"
FLUSSO_POP_RIC = "IT1,164_164_DF_DCIS_RICPOPRES2011_1,1.0"
ANNO_ISTAT_INIZIO = 2000
VARESE, REGIONE, ITALIA = "ITC41", "ITC4", "IT"
PROVINCE_LOMBARDE = ["ITC41", "ITC42", "ITC43", "ITC44", "ITC45", "ITC46", "ITC47", "ITC48", "ITC49", "ITC4A", "ITC4B", "IT108"]
NOMI_LOMBARDI = ["Varese", "Como", "Sondrio", "Milano", "Bergamo", "Brescia", "Pavia", "Cremona", "Mantova", "Lecco", "Lodi",
                 "Monza e della Brianza"]
NOMI_OBBLIGATORI = NOMI_LOMBARDI + ["Lombardia", "Italia"]
RIPARTIZIONI = ["Nord-ovest", "Nord-est", "Centro", "Sud e Isole"]
COLONNE_TAGLIACARNE = ["agricoltura", "industria", "costruzioni", "commercio_trasporti_ict", "finanza_immobiliare_professioni",
                       "altri_servizi", "totale", "procapite"]
INTESTAZIONI_ATTESE = ["provincia", "agricoltura", "industria in senso stretto", "costruzioni", "commercio", "attivita finanziarie",
                       "altri servizi", "totale", "valore aggiunto procapite"]
CSV_TAGLIACARNE = ["anno", "territorio"] + COLONNE_TAGLIACARNE + ["fonte"]
CSV_ISTAT = ["anno", "area", "valore_aggiunto", "popolazione_media"]
SOGLIA_AVVISO = 0.08      # variazione annua oltre ±8%: avviso
SOGLIA_BLOCCO = 0.15      # oltre ±15%: blocco (probabilmente l'anno prima è stato rivisto)
SCARTO_AVVISO = 0.04      # crescita diversa da quella dell'Italia di oltre 4 punti: avviso
SCARTO_BLOCCO = 0.10      # oltre 10 punti: blocco
PAUSA_ISTAT = 13          # ISTAT accetta circa 5 richieste al minuto per indirizzo
PAUSE_SITO = [30, 90, 180]
PAUSE_ISTAT_TENTATIVI = [60, 180, 300]
TIMEOUT = 300             # la richiesta del valore aggiunto di tutte le province pesa ~800 KB e da GitHub ha impiegato fino a 80 secondi
GIORNI_MAX_ISTAT = 28
USER_AGENT = "Mozilla/5.0 (report-conti-economici; Camera di Commercio di Varese)"


class Problema(Exception):
    """tipo: "rete" (il sito non risponde), "risposta" (risponde ma non con quello che serve), "formato" (il file o la
    tavola sono cambiati), "anomalia" (numeri non plausibili o diversi da quelli già salvati), "prova" (simulato)."""
    def __init__(self, tipo, dettaglio):
        super().__init__(dettaglio)
        self.tipo = tipo


# ---------------------------------------------------------------------------------------------------- rete
def scarica(url, accept, pause, nome_sito):
    errore = None
    for n in range(len(pause) + 1):
        try:
            req = urllib.request.Request(url, headers={"Accept": accept, "Accept-Language": "it, en;q=0.5", "User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return r.read()
        except Exception as e:  # rete, 429, 5xx, timeout
            errore = e
            print(f"  tentativo {n + 1} non riuscito: {e}", flush=True)
            if n < len(pause):
                time.sleep(pause[n])
    n = len(pause) + 1
    raise Problema("rete", f"il sito {nome_sito} non ha risposto ({n} {'tentativo' if n == 1 else 'tentativi'}): {errore}")


# ---------------------------------------------------------------------------------------------------- Excel (solo libreria standard)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def _colonna(rif):
    n = 0
    for c in re.match(r"[A-Z]+", rif).group(0):
        n = n * 26 + ord(c) - 64
    return n - 1


def leggi_xlsx(dati):
    """Contenuto di un .xlsx -> {nome del foglio: lista di righe (liste di valori: testo, numero o None)}."""
    try:
        z = zipfile.ZipFile(io.BytesIO(dati))
        z.testzip()
    except Exception as e:
        raise Problema("risposta", f"il file scaricato non è un Excel (.xlsx) valido: {e}")
    try:
        condivise = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
                condivise.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        fogli = {}
        for s in wb.find("m:sheets", NS):
            bers = rel[s.get("{%s}id" % NS["r"])]
            percorso = bers.lstrip("/") if bers.startswith("/") else "xl/" + bers
            righe = []
            for r in ET.fromstring(z.read(percorso)).find("m:sheetData", NS):
                riga = {}
                for c in r:
                    t, v = c.get("t"), c.find("m:v", NS)
                    if t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
                    elif v is None or v.text is None:
                        val = None
                    elif t == "s":
                        val = condivise[int(v.text)]
                    elif t in ("str", "e"):
                        val = v.text
                    elif t == "b":
                        val = bool(int(v.text))
                    else:
                        val = float(v.text)
                    riga[_colonna(c.get("r"))] = val
                n = r.get("r")
                righe.append((int(n) if n else len(righe) + 1, riga))
            massimo = max((k for _, r in righe for k in r), default=-1)
            tabella = []
            for numero, riga in righe:
                while len(tabella) < numero - 1:
                    tabella.append([None] * (massimo + 1))
                tabella.append([riga.get(j) for j in range(massimo + 1)])
            fogli[s.get("name")] = tabella
        return fogli
    except Problema:
        raise
    except Exception as e:
        raise Problema("formato", f"non riesco a leggere la struttura del file Excel: {type(e).__name__}: {e}")


# ---------------------------------------------------------------------------------------------------- TAGLIACARNE
def norm(s):
    s = html.unescape(str(s)).lower().replace("à", "a").replace("è", "e").replace("é", "e").replace("ì", "i").replace("ò", "o").replace("ù", "u")
    return re.sub(r"\s+", " ", s).strip()


def trova_file(pagina_html):
    """Pagina Tagliacarne -> (url del file, anno, versione gg-mm-aaaa o None, testo del link)."""
    trovati = []
    for m in re.finditer(r"<a\b[^>]*?href=\"([^\"]+\.xlsx)\"[^>]*>(.*?)</a>", pagina_html, re.S | re.I):
        testo = re.sub(r"<[^>]+>", "", html.unescape(m.group(2)))
        testo = re.sub(r"\s+", " ", testo.replace("\xa0", " ")).strip()
        t = norm(testo)
        if t.startswith("valore aggiunto per provincia e branca"):
            a = re.search(r"anno (\d{4})", t)
            v = re.search(r"versione (?:del )?(\d{1,2}-\d{1,2}-\d{4})", t)
            if a:
                trovati.append((urllib.parse.urljoin(PAGINA_TAGLIACARNE, html.unescape(m.group(1))), int(a.group(1)), v.group(1) if v else None, testo))
    if not trovati:
        raise Problema("formato", "nella pagina del Centro Studi Tagliacarne non trovo più il link «Valore aggiunto per provincia e branca di "
                                  "attività economica. Anno …» (la pagina è stata cambiata o il file è stato tolto)")
    trovati.sort(key=lambda x: x[1])
    return trovati[-1]


def leggi_tavola_tagliacarne(fogli):
    """Fogli Excel -> (anno, {territorio: [8 valori]}). Controlla intestazioni e righe attese."""
    candidati = []
    for nome, righe in fogli.items():
        titolo = norm(righe[0][0]) if righe and righe[0] and righe[0][0] else ""
        a = re.search(r"anno (\d{4})", titolo) or re.fullmatch(r"(\d{4})", norm(nome))
        if a and righe and len(righe) > 4:
            candidati.append((int(a.group(1)), nome, righe))
    if not candidati:
        raise Problema("formato", "nel file Excel non c'è nessun foglio con il titolo «… Anno AAAA» (il file è cambiato)")
    anno, nome, righe = sorted(candidati, key=lambda x: x[0])[-1]
    testa = None
    for i, r in enumerate(righe[:8]):
        if r and r[0] is not None and norm(r[0]) == "provincia":
            testa = i
            break
    if testa is None:
        raise Problema("formato", f"nel foglio «{nome}» non trovo la riga di intestazione che comincia con «Provincia»")
    for j, atteso in enumerate(INTESTAZIONI_ATTESE):
        h = norm(righe[testa][j]) if len(righe[testa]) > j and righe[testa][j] is not None else ""
        if not h.startswith(atteso):
            raise Problema("formato", f"colonna {j + 1} del file Excel: trovo «{righe[testa][j] if len(righe[testa]) > j else ''}» invece di «{atteso}» (il file è cambiato)")
    if len(righe[testa]) > len(INTESTAZIONI_ATTESE) and any(x is not None for x in righe[testa][len(INTESTAZIONI_ATTESE):]):
        raise Problema("formato", "il file Excel ha colonne in più rispetto a quelle attese (il file è cambiato)")
    out = {}
    for r in righe[testa + 1:]:
        if not r or r[0] is None or str(r[0]).strip() == "":
            continue
        nomeT = re.sub(r"\s+", " ", str(r[0])).strip()
        v = r[1:9]
        if len(v) < 8 or any(not isinstance(x, (int, float)) or isinstance(x, bool) for x in v):
            if all(x is None for x in v):
                continue                        # nota o riga di commento senza numeri
            raise Problema("formato", f"riga «{nomeT}» del file Excel: valori mancanti o non numerici")
        v = [float(x) for x in v]
        if nomeT in out:
            # una provincia che è anche regione (Valle d'Aosta) compare due volte con gli stessi numeri: va bene; con numeri diversi no
            if any(abs(x - y) > 0.01 for x, y in zip(out[nomeT], v)):
                raise Problema("formato", f"la riga «{nomeT}» compare due volte nel file Excel con numeri diversi")
            continue
        out[nomeT] = v
    mancano = [n for n in NOMI_OBBLIGATORI if n not in out]
    if mancano:
        raise Problema("formato", f"nel file Excel mancano le righe {', '.join(mancano)} (nomi cambiati o righe tolte)")
    if len(out) < 120:
        raise Problema("formato", f"il file Excel ha {len(out)} righe di territori invece di circa 132 (righe tolte?)")
    return anno, out


def controlla_tavola_tagliacarne(anno, T):
    """Controlli di coerenza interna del file (bloccanti)."""
    for n, v in T.items():
        if abs(sum(v[:6]) - v[6]) > 0.35:     # tolleranza: nei file con valori già arrotondati a 1 decimale la somma può scostarsi fino a 0,3
            raise Problema("anomalia", f"{n}: la somma delle branche ({sum(v[:6]):.1f}) non coincide con il totale ({v[6]:.1f})")
        if not (v[6] > 0 and 5000 <= v[7] <= 150000):
            raise Problema("anomalia", f"{n}: totale o valore pro-capite non plausibili ({v[6]:.1f}; {v[7]:.1f})")
    somma = sum(T[n][6] for n in NOMI_LOMBARDI)
    if abs(somma - T["Lombardia"][6]) > 0.5:
        raise Problema("anomalia", f"la somma delle 12 province lombarde ({somma:.1f}) non coincide con la Lombardia ({T['Lombardia'][6]:.1f})")
    mac = [n for n in RIPARTIZIONI if n in T]
    if len(mac) == len(RIPARTIZIONI):
        sm = sum(T[n][6] for n in mac)
        extra = T["Italia"][6] - sm
        if not (0 <= extra <= 0.003 * T["Italia"][6]):
            raise Problema("anomalia", f"Italia ({T['Italia'][6]:.1f}) e somma delle ripartizioni ({sm:.1f}): lo scarto ({extra:.1f}) non è quello atteso "
                                       "(l'«extra-regio» è di poche centinaia di milioni)")


def csv_tagliacarne(anno, T, fonte):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(CSV_TAGLIACARNE)
    for n, v in T.items():
        w.writerow([anno, n] + [repr(x) for x in v] + [fonte])
    return buf.getvalue()


def leggi_csv_tagliacarne(percorso):
    T, anno, fonte = {}, None, None
    with open(percorso, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            anno = int(r["anno"])
            fonte = r.get("fonte")
            T[r["territorio"]] = [float(r[c]) if r[c] != "" else None for c in COLONNE_TAGLIACARNE]
    return anno, T, fonte


def archivio_tagliacarne():
    """{anno: nome del file} dei file già salvati."""
    out = {}
    if os.path.isdir(CARTELLA):
        for n in os.listdir(CARTELLA):
            m = re.fullmatch(r"tagliacarne_(\d{4})\.csv", n)
            if m:
                out[int(m.group(1))] = n
    return out


def confronto_anno_prima(anno, T, archivio, avvisi, accetta):
    """Confronta ogni territorio con l'archivio dell'anno prima. Segnala una base probabilmente rivista (cioè un anno prima
    diverso da quello su cui Tagliacarne ha calcolato la crescita) in due modi:
      - variazione annua assoluta oltre ±8% (avviso) o ±15% (blocco);
      - scarto della crescita rispetto a quella dell'Italia oltre 4 punti (avviso) o 10 punti (blocco): le province crescono in modo simile
        (nel 2024 da -0,2% a +4,9% con l'Italia a +2,1%), mentre una base rivista sposta una provincia da sola
        (esempio vero: Varese 2023 era 27.800 nell'edizione del 2024 e 29.091 in quella del 2025, +4,6% a parità di Italia +0,8%)."""
    prima = archivio.get(anno - 1)
    if not prima:
        avvisi.append(f"nell'archivio non c'è il file Tagliacarne del {anno - 1}: la pagina non potrà fare il confronto con l'anno prima")
        return
    _, P, _ = leggi_csv_tagliacarne(os.path.join(CARTELLA, prima))
    var = {n: T[n][6] / P[n][6] - 1 for n in T if n in P and P[n][6]}
    ita = var.get("Italia")
    fuori, scarti = [], []
    for n, g in var.items():
        if abs(g) > SOGLIA_AVVISO:
            fuori.append((n, g))
        if ita is not None and n != "Italia" and abs(g - ita) > SCARTO_AVVISO:
            scarti.append((n, g - ita))
    gravi = [x for x in fuori if abs(x[1]) > SOGLIA_BLOCCO] + [x for x in scarti if abs(x[1]) > SCARTO_BLOCCO]
    if gravi and not accetta:
        dettaglio = ", ".join(f"{n} {g * 100:+.1f}%" for n, g in fuori) or ", ".join(f"{n} {g * 100:+.1f} punti rispetto all'Italia" for n, g in scarti)
        raise Problema("anomalia", f"variazione {anno}/{anno - 1} fuori dalla norma ({dettaglio}): probabilmente Tagliacarne ha rivisto i dati del {anno - 1}, "
                                   f"diversi da quelli archiviati. Dati NON salvati: verificare con il comunicato stampa di Tagliacarne e, se giusto, "
                                   "rilanciare con «Accetta»")
    if fuori:
        avvisi.append(f"variazioni {anno}/{anno - 1} sopra il {int(SOGLIA_AVVISO * 100)}%: " + ", ".join(f"{n} {g * 100:+.1f}%" for n, g in fuori))
    if scarti:
        avvisi.append(f"crescita {anno}/{anno - 1} molto diversa da quella dell'Italia ({ita * 100:+.1f}%): " +
                      ", ".join(f"{n} {g * 100 + ita * 100:+.1f}%" for n, g in scarti) + " (la base del " + str(anno - 1) + " potrebbe essere stata rivista)")


def parte_tagliacarne(info, args, esiti, problemi, avvisi):
    nome = "tagliacarne"
    salvato = info.get(nome, {})
    print("Tagliacarne: leggo la pagina del Centro Studi…", flush=True)
    pagina = scarica(PAGINA_TAGLIACARNE, "text/html", PAUSE_SITO, "del Centro Studi Tagliacarne").decode("utf-8", "replace")
    url, anno_link, versione, testo_link = trova_file(pagina)
    print(f"Tagliacarne: file «{testo_link}» → {url}", flush=True)
    dati = scarica(url, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*", PAUSE_SITO, "del Centro Studi Tagliacarne")
    sha = hashlib.sha256(dati).hexdigest()
    archivio = archivio_tagliacarne()
    percorso_anno = os.path.join(CARTELLA, f"tagliacarne_{anno_link}.csv")
    if (not args.forza and salvato.get("sha256") == sha and salvato.get("anno") == anno_link and os.path.exists(percorso_anno)):
        salvato.update({"controllato": ora()})
        info[nome] = salvato
        esiti.append({"parte": nome, "novita": False, "anno": anno_link})
        print("Tagliacarne: nessuna novità (stesso file già salvato).", flush=True)
        return
    anno, T = leggi_tavola_tagliacarne(leggi_xlsx(dati))
    if anno != anno_link:
        raise Problema("formato", f"il link dice «Anno {anno_link}» ma dentro il file c'è il {anno}")
    ultimo_salvato = max(archivio) if archivio else None
    if ultimo_salvato and anno < ultimo_salvato:
        raise Problema("anomalia", f"il file Tagliacarne in linea è del {anno}, ma l'archivio ha già il {ultimo_salvato}: tenuto l'archivio")
    controlla_tavola_tagliacarne(anno, T)
    confronto_anno_prima(anno, T, archivio, avvisi, args.accetta)
    nuovo = csv_tagliacarne(anno, T, "tagliacarne")
    modificato = True
    if anno in archivio:
        a0, T0, f0 = leggi_csv_tagliacarne(percorso_anno)
        diff = [n for n in T if n in T0 and T0[n][6] and abs(T[n][6] - T0[n][6]) > 0.05]     # correzioni di almeno 50.000 euro
        grandi = [n for n in diff if abs(T[n][6] / T0[n][6] - 1) > 0.005]
        if grandi and not args.accetta:
            raise Problema("anomalia", f"il file Tagliacarne del {anno} è stato ricaricato con valori diversi da quelli già salvati "
                                       f"({len(diff)} territori cambiati, ad esempio {', '.join(grandi[:4])}): dati NON sostituiti. "
                                       "Verificare e, se giusto, rilanciare con «Accetta»")
        if diff:
            avvisi.append(f"il file Tagliacarne del {anno} è stato ricaricato con piccole correzioni ({len(diff)} territori): sostituito")
        elif set(T) == set(T0) and all(abs(x - y) <= 1e-6 * max(1.0, abs(y)) for n in T for x, y in zip(T[n], T0[n]) if y is not None):
            modificato = False           # stessi numeri (cambia al massimo l'ultima cifra decimale, per un file salvato di nuovo con Excel)
    if modificato:
        with open(percorso_anno + ".tmp", "w", encoding="utf-8", newline="") as f:
            f.write(nuovo)
        os.replace(percorso_anno + ".tmp", percorso_anno)
    info[nome] = {"anno": anno, "versione": versione, "url": url, "titolo": testo_link, "sha256": sha, "dimensione": len(dati), "righe": len(T),
                  "scaricato": ora(), "controllato": ora(),
                  "dati_modificati": ora() if modificato else salvato.get("dati_modificati", ora())}
    esiti.append({"parte": nome, "novita": modificato, "anno": anno, "versione": versione})
    print(f"Tagliacarne: anno {anno} (versione {versione}), {len(T)} territori, {'SALVATO' if modificato else 'invariato'}", flush=True)


# ---------------------------------------------------------------------------------------------------- ISTAT
def url_dati(flusso, chiave, inizio):
    return f"{SDMXWS}data/{flusso}/{chiave}?startPeriod={inizio}&format=csvfile"


def ultimo_aggiornamento(flusso):
    """Data di ultimo aggiornamento dichiarata da ISTAT per il dataflow (facoltativa: se non c'è, None)."""
    try:
        xml = scarica(f"{SDMXWS}dataflow/{flusso.replace(',', '/')}?references=none", "application/xml", [], "ISTAT").decode("utf-8", "replace")
        i = xml.find('id="LAST_UPDATE"')
        a = xml.find("<common:AnnotationTitle>", i)
        b = xml.find("</common:AnnotationTitle>", a)
        return xml[a + 24:b] if i > 0 and a > 0 and b > a else None
    except Exception:
        return None


def leggi_csv_istat(dati, richieste, nome):
    testo = dati.decode("utf-8-sig", "replace")
    prima = testo.split("\n", 1)[0]
    if not all(c in prima for c in richieste):
        raise Problema("risposta", f"ISTAT ha risposto, ma non con il CSV dei dati «{nome}» (inizio: «{testo[:120].strip()[:120]}»)")
    return list(csv.DictReader(io.StringIO(testo)))


def chiave_edizione(e):
    m = re.fullmatch(r"(\d{4})M(\d{1,2})", e or "")
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def serie_istat_va(righe):
    """Righe ISTAT -> ({area: {anno: valore}}, edizione più recente). Per ogni area e anno vale l'edizione più recente."""
    best = {}
    for r in righe:
        if r["OBS_VALUE"] in ("", None):
            continue
        try:
            anno, val = int(r["TIME_PERIOD"]), float(r["OBS_VALUE"])
        except ValueError:
            raise Problema("formato", f"valore o periodo non numerico nel CSV ISTAT (periodo «{r['TIME_PERIOD']}», valore «{r['OBS_VALUE']}»)")
        k = (r["REF_AREA"], anno)
        e = chiave_edizione(r.get("EDITION"))
        if k not in best or e > best[k][0]:
            best[k] = (e, val, r.get("EDITION"))
    aree = {}
    for (a, y), (_, v, _) in best.items():
        aree.setdefault(a, {})[y] = v
    ed = max((e for e, _, _ in best.values()), default=(0, 0))
    return aree, f"{ed[0]}M{ed[1]}"


def elenco_province(aree):
    return [a for a in aree if (len(a) == 5 and a.startswith("IT") and a != "ITCDE") or a in ("IT108", "IT109", "IT110", "IT111")]


def calcola_serie_va(aree):
    """-> {area: {anno: valore arrotondato a 1 decimale}} per VARESE, LOMBARDIA, ITALIA; anni completi."""
    out = {"VARESE": {}, "LOMBARDIA": {}, "ITALIA": {}}
    if VARESE not in aree or any(p not in aree for p in PROVINCE_LOMBARDE):
        raise Problema("formato", "nei dati ISTAT mancano Varese o alcune province lombarde (codici cambiati?)")
    prov = elenco_province(aree)
    if len(prov) < 100 or "ITZ" not in aree or ITALIA not in aree:
        raise Problema("formato", f"nei dati ISTAT trovo {len(prov)} province (attese circa 107), «extra-regio» (ITZ) o l'Italia (IT) mancano")
    for y in sorted(aree[VARESE]):
        if y < ANNO_ISTAT_INIZIO:
            continue
        if not all(y in aree[p] for p in prov) or y not in aree["ITZ"]:
            continue                                              # anno non ancora completo per tutte le province
        lom = sum(aree[p][y] for p in PROVINCE_LOMBARDE)
        ita = sum(aree[p][y] for p in prov) + aree["ITZ"][y]
        if y in aree[ITALIA] and abs(ita - aree[ITALIA][y]) > 0.0001 * aree[ITALIA][y]:
            raise Problema("anomalia", f"{y}: somma delle province ISTAT ({ita:.1f}) diversa dal dato nazionale ({aree[ITALIA][y]:.1f}): provincia mancante o codice cambiato?")
        if y in aree[REGIONE] and abs(lom - aree[REGIONE][y]) > 0.0005 * aree[REGIONE][y]:
            raise Problema("anomalia", f"{y}: somma delle 12 province lombarde ({lom:.1f}) diversa dalla Lombardia ISTAT ({aree[REGIONE][y]:.1f})")
        out["VARESE"][y] = round_half_up(aree[VARESE][y], 1)
        out["LOMBARDIA"][y] = round_half_up(lom, 1)
        out["ITALIA"][y] = round_half_up(ita, 1)
    if not out["VARESE"]:
        raise Problema("formato", "nei dati ISTAT non c'è nessun anno completo")
    return out


def round_half_up(x, n):
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(format(x, ".15g")).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP))


def serie_popolazione(righe_attuali, righe_ric):
    pop = {"VARESE": {}, "LOMBARDIA": {}, "ITALIA": {}}
    codici = {VARESE: "VARESE", REGIONE: "LOMBARDIA", ITALIA: "ITALIA"}
    for righe in (righe_attuali, righe_ric):          # la serie attuale ha la precedenza sulla ricostruzione
        for r in righe:
            if r["REF_AREA"] in codici and r["OBS_VALUE"] not in ("", None):
                pop[codici[r["REF_AREA"]]].setdefault(int(r["TIME_PERIOD"]), float(r["OBS_VALUE"]))
    return pop


def csv_istat(va, pop):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(CSV_ISTAT)
    for a in ("VARESE", "LOMBARDIA", "ITALIA"):
        for y in sorted(va[a]):
            media = (pop[a][y] + pop[a][y + 1]) / 2 if y in pop[a] and y + 1 in pop[a] else None
            w.writerow([y, a, repr(va[a][y]), "" if media is None else repr(media)])
    return buf.getvalue()


def leggi_csv_istat_salvato(percorso):
    va = {"VARESE": {}, "LOMBARDIA": {}, "ITALIA": {}}
    if os.path.exists(percorso):
        with open(percorso, encoding="utf-8") as f:
            for r in csv.DictReader(f):
                va[r["area"]][int(r["anno"])] = float(r["valore_aggiunto"])
    return va


def parte_istat(info, args, esiti, problemi, avvisi):
    nome = "istat_valore_aggiunto"
    salvato = info.get(nome, {})
    percorso = os.path.join(CARTELLA, nome + ".csv")
    print("ISTAT: controllo la data di aggiornamento…", flush=True)
    agg = ultimo_aggiornamento(FLUSSO_VA)
    ultimo = salvato.get("scaricato")
    vecchio = True
    if ultimo:
        vecchio = (datetime.now(timezone.utc) - datetime.strptime(ultimo, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)).days >= GIORNI_MAX_ISTAT
    if not (args.forza or not os.path.exists(percorso) or agg is None or agg != salvato.get("ultimo_aggiornamento_istat") or vecchio):
        salvato["controllato"] = ora()
        info[nome] = salvato
        esiti.append({"parte": nome, "novita": False, "ultimo_anno": salvato.get("ultimo_anno")})
        print(f"ISTAT: nessuna novità (ultimo aggiornamento {agg}).", flush=True)
        return
    print("ISTAT: scarico valore aggiunto e popolazione…", flush=True)
    time.sleep(PAUSA_ISTAT)
    ris_va = leggi_csv_istat(scarica(url_dati(FLUSSO_VA, "A..B1G_B_W2_S1.V..", ANNO_ISTAT_INIZIO), "text/csv", PAUSE_ISTAT_TENTATIVI, "ISTAT"),
                             ["REF_AREA", "TIME_PERIOD", "OBS_VALUE", "EDITION"], "valore aggiunto provinciale")
    time.sleep(PAUSA_ISTAT)
    pop1 = leggi_csv_istat(scarica(url_dati(FLUSSO_POP, f"A.{VARESE}+{REGIONE}+{ITALIA}.JAN.9.TOTAL.99", ANNO_ISTAT_INIZIO), "text/csv", PAUSE_ISTAT_TENTATIVI, "ISTAT"),
                           ["REF_AREA", "TIME_PERIOD", "OBS_VALUE"], "popolazione residente")
    time.sleep(PAUSA_ISTAT)
    pop2 = leggi_csv_istat(scarica(url_dati(FLUSSO_POP_RIC, f"A.{VARESE}+{REGIONE}+{ITALIA}.JAN.TOTAL.9.TOTAL", ANNO_ISTAT_INIZIO), "text/csv", PAUSE_ISTAT_TENTATIVI, "ISTAT"),
                           ["REF_AREA", "TIME_PERIOD", "OBS_VALUE"], "popolazione ricostruita")
    aree, edizione = serie_istat_va(ris_va)
    va = calcola_serie_va(aree)
    pop = serie_popolazione(pop1, pop2)
    for a in va:
        if not pop[a]:
            raise Problema("formato", f"nei dati ISTAT sulla popolazione manca {a}")
    vecchi = leggi_csv_istat_salvato(percorso)
    ultimo_nuovo, ultimo_vecchio = max(va["VARESE"]), max(vecchi["VARESE"]) if vecchi["VARESE"] else None
    if ultimo_vecchio and ultimo_nuovo < ultimo_vecchio:
        raise Problema("anomalia", f"i dati ISTAT arrivano al {ultimo_nuovo}, quelli già salvati al {ultimo_vecchio}: tenuti quelli salvati")
    n_nuovo, n_vecchio = sum(len(v) for v in va.values()), sum(len(v) for v in vecchi.values())
    if n_vecchio and n_nuovo < 0.9 * n_vecchio:
        raise Problema("anomalia", f"i dati ISTAT hanno {n_nuovo} valori, quelli già salvati {n_vecchio}: tenuti quelli salvati")
    cambi = [(a, y, vecchi[a][y], va[a][y]) for a in va for y in va[a] if y in vecchi[a] and abs(va[a][y] - vecchi[a][y]) > 0.05]
    grandi = [c for c in cambi if abs(c[3] / c[2] - 1) > 0.10]
    if grandi and not args.accetta:
        raise Problema("anomalia", "ISTAT ha cambiato di più del 10% alcuni valori già salvati (" +
                       ", ".join(f"{a} {y}: {o:.1f} → {n:.1f}" for a, y, o, n in grandi[:4]) + "): dati NON sostituiti. Verificare e, se giusto, rilanciare con «Accetta»")
    if cambi:
        avvisi.append(f"ISTAT ha rivisto {len(cambi)} valori già salvati (variazione massima {max(abs(c[3] / c[2] - 1) for c in cambi) * 100:.2f}%)")
    # confronto informativo con l'archivio Tagliacarne: dove i due metodi si incontrano
    for anno_t, file_t in sorted(archivio_tagliacarne().items()):
        _, Tt, _ = leggi_csv_tagliacarne(os.path.join(CARTELLA, file_t))
        for a, nm in (("VARESE", "Varese"), ("LOMBARDIA", "Lombardia"), ("ITALIA", "Italia")):
            if anno_t in va[a] and nm in Tt and Tt[nm][6]:
                avvisi.append(f"confronto ISTAT / Tagliacarne {anno_t}, {nm}: {va[a][anno_t]:.1f} / {Tt[nm][6]:.1f} ({(va[a][anno_t] / Tt[nm][6] - 1) * 100:+.2f}%)")
    nuovo = csv_istat(va, pop)
    modificato = not os.path.exists(percorso) or open(percorso, encoding="utf-8").read() != nuovo
    if modificato:
        with open(percorso + ".tmp", "w", encoding="utf-8", newline="") as f:
            f.write(nuovo)
        os.replace(percorso + ".tmp", percorso)
    info[nome] = {"scaricato": ora(), "controllato": ora(), "ultimo_anno": ultimo_nuovo, "righe": n_nuovo, "edizione": edizione,
                  "ultimo_aggiornamento_istat": agg, "dati_modificati": ora() if modificato else salvato.get("dati_modificati", ora())}
    esiti.append({"parte": nome, "novita": modificato, "ultimo_anno": ultimo_nuovo, "prima": ultimo_vecchio})
    print(f"ISTAT: {n_nuovo} valori, ultimo anno {ultimo_nuovo}, edizione {edizione}, ultimo aggiornamento {agg}, {'MODIFICATO' if modificato else 'invariato'}", flush=True)


# ---------------------------------------------------------------------------------------------------- principale
def ora():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def main(argv=None):
    arg = argparse.ArgumentParser(description="Aggiorna i dati del valore aggiunto (Tagliacarne + ISTAT)")
    arg.add_argument("--esito", help="file JSON in cui scrivere l'esito (lo legge segnala_problemi.py)")
    arg.add_argument("--forza", action="store_true", help="scarica e ricontrolla tutto anche senza novità")
    arg.add_argument("--accetta", action="store_true", help="salva anche se un controllo di plausibilità segnala un'anomalia")
    args = arg.parse_args(argv)
    prova = os.environ.get("PROVA_ERRORE", "").lower() == "true"
    os.makedirs(CARTELLA, exist_ok=True)
    p_info = os.path.join(CARTELLA, "aggiornamento.json")
    info = json.load(open(p_info, encoding="utf-8")) if os.path.exists(p_info) else {}
    problemi, esiti, avvisi = [], [], []
    parti = [("tagliacarne", "Valore aggiunto del Centro Studi Tagliacarne", parte_tagliacarne),
             ("istat_valore_aggiunto", "Valore aggiunto provinciale ISTAT", parte_istat)]
    for cod, nome, fn in parti:
        try:
            if prova and cod == parti[-1][0]:
                raise Problema("prova", "errore simulato per provare l'e-mail di avviso (nessun problema reale)")
            fn(info, args, esiti, problemi, avvisi)
        except Problema as e:
            problemi.append({"tavola": cod, "nome": nome, "tipo": e.tipo, "dettaglio": str(e), "ultimo_anno_salvato": info.get(cod, {}).get("anno") or info.get(cod, {}).get("ultimo_anno")})
            print(f"{cod}: PROBLEMA ({e.tipo}) {e}", flush=True)
        except Exception as e:  # imprevisto: lo si segnala come tale
            problemi.append({"tavola": cod, "nome": nome, "tipo": "imprevisto", "dettaglio": f"{type(e).__name__}: {e}",
                             "ultimo_anno_salvato": info.get(cod, {}).get("anno") or info.get(cod, {}).get("ultimo_anno")})
            print(f"{cod}: ERRORE IMPREVISTO {type(e).__name__}: {e}", flush=True)
    info["controllato"] = ora()
    info["avvisi"] = avvisi
    info["archivio_tagliacarne"] = sorted(archivio_tagliacarne())      # gli anni dei file dati/tagliacarne_AAAA.csv (la pagina li legge da qui)
    with open(p_info, "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=1)
        f.write("\n")
    if esito_ok(args):
        with open(args.esito, "w", encoding="utf-8") as f:
            json.dump({"quando": info["controllato"], "prova": prova, "problemi": problemi,
                       "tavole": [{"tavola": e["parte"], "nome": e["parte"], "ultimo_anno": e.get("anno") or e.get("ultimo_anno"),
                                   "cambiato": e.get("novita")} for e in esiti], "avvisi": avvisi}, f, ensure_ascii=False, indent=1)
    for a in avvisi:
        print("AVVISO:", a, flush=True)
    if problemi:
        print("\n".join(f"{p['tavola']}: {p['dettaglio']}" for p in problemi), file=sys.stderr)
        if not args.esito:
            sys.exit(1)


def esito_ok(args):
    return bool(args.esito)


if __name__ == "__main__":
    main()
