# -*- coding: utf-8 -*-
"""
GALLEON REPORT - pipeline automatica (gira in locale: nessun dato esce dal PC)
==============================================================================
Ogni esecuzione (lanciata ogni notte dall'Utilità di pianificazione di Windows):
  1) (opzionale) scarica gli allegati dalle email configurate -> dati/inbox/...
  2) importa i file NUOVI di dati/inbox nel database locale SQLite e li sposta
     in dati/archivio (i file illeggibili finiscono in dati/scarti + log)
  3) ricalcola tutto: conto economico, KPI, menu engineering, consumi/sprechi
  4) scrive i report pronti:  report/Galleon_report.html  e  report/Galleon_report.xlsx
     (+ copia datata in report/storico, + nota cruscotto in Obsidian se configurata)

È idempotente: se una chiusura, una fattura o un file presenze viene
riesportato, i dati di quel giorno/documento vengono SOSTITUITI, mai duplicati.

Uso:   python pipeline.py
"""
import datetime as dt
import email
import hashlib
import imaplib
import json
import logging
import os
import re
import shutil
import sqlite3
import xml.etree.ElementTree as ET
from html import escape
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# 0. CONFIGURAZIONE E CARTELLE
# --------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
CFG = json.loads((BASE / "config.json").read_text(encoding="utf-8"))


def cartella(nome):
    p = BASE / CFG["cartelle"][nome]
    p.mkdir(parents=True, exist_ok=True)
    return p


INBOX, ARCHIVIO, SCARTI = cartella("inbox"), cartella("archivio"), cartella("scarti")
ANAG, REPORT, LOGDIR = cartella("anagrafiche"), cartella("report"), cartella("log")
DB = BASE / CFG["database"]
DB.parent.mkdir(parents=True, exist_ok=True)
SEP, DEC = CFG["csv"]["separatore"], CFG["csv"]["decimale"]
OBIETTIVI = CFG["obiettivi"]

logging.basicConfig(filename=LOGDIR / "pipeline.log", level=logging.INFO, encoding="utf-8",
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("galleon")
log.addHandler(logging.StreamHandler())

SCHEMA = """
CREATE TABLE IF NOT EXISTS file_importati (hash TEXT PRIMARY KEY, nome TEXT, quando TEXT);
CREATE TABLE IF NOT EXISTS vendite  (data TEXT, canale TEXT, codice TEXT, descrizione TEXT, categoria TEXT,
                                     quantita REAL, importo_lordo REAL, scontrinato TEXT);
CREATE TABLE IF NOT EXISTS coperti  (data TEXT, canale TEXT, coperti REAL);
CREATE TABLE IF NOT EXISTS presenze (data TEXT, dipendente TEXT, ore REAL, reparto TEXT);
CREATE TABLE IF NOT EXISTS inventari(data TEXT, ingrediente TEXT, quantita REAL);
CREATE TABLE IF NOT EXISTS fatture_righe (piva TEXT, fornitore TEXT, numero TEXT, data TEXT, tipo TEXT,
                                     codice_art TEXT, descrizione TEXT, quantita REAL, um TEXT,
                                     importo REAL, aliquota REAL);
"""

# Colonne del FORMATO STANDARD interno. Gli export reali vengono adattati a questo.
COLONNE = {
    "vendite": ["data", "canale", "codice", "descrizione", "categoria", "quantita", "importo_lordo", "scontrinato"],
    "coperti": ["data", "canale", "coperti"],
    "presenze": ["data", "dipendente", "ore", "reparto"],
}


# --------------------------------------------------------------------------
# 1. RACCOLTA AUTOMATICA DA EMAIL (opzionale, attiva in config.json)
# --------------------------------------------------------------------------
def scarica_allegati_email():
    """Legge le email NON LETTE e salva gli allegati in dati/inbox/<cartella> secondo le regole.
    La password NON sta nel file: si legge da una variabile d'ambiente di Windows."""
    cfg = CFG.get("email", {})
    if not cfg.get("attivo"):
        return
    pwd = os.environ.get(cfg["variabile_password"], "")
    with imaplib.IMAP4_SSL(cfg["server"]) as M:
        M.login(cfg["utente"], pwd)
        M.select(cfg.get("casella", "INBOX"))
        _, ids = M.search(None, "UNSEEN")
        for i in ids[0].split():
            _, dati = M.fetch(i, "(RFC822)")
            msg = email.message_from_bytes(dati[0][1])
            mitt, ogg = (msg.get("From") or "").lower(), (msg.get("Subject") or "").lower()
            for r in cfg["regole"]:
                if r["mittente_contiene"].lower() in mitt and r.get("oggetto_contiene", "").lower() in ogg:
                    for parte in msg.walk():  # walk entra anche nelle email annidate (PEC/SDI)
                        nome = parte.get_filename()
                        if not nome or nome.lower().startswith("daticert") or "_mt_" in nome.lower():
                            continue
                        if not any(nome.lower().endswith(e) for e in r["estensioni"]):
                            continue
                        dest = INBOX / r["cartella"]
                        dest.mkdir(exist_ok=True)
                        (dest / Path(nome).name).write_bytes(parte.get_payload(decode=True))
                        log.info(f"email -> {r['cartella']}/{nome}")
                    break


# --------------------------------------------------------------------------
# 2. IMPORTAZIONE FILE -> DATABASE LOCALE
# --------------------------------------------------------------------------
def leggi_csv(f):
    df = pd.read_csv(f, sep=SEP, decimal=DEC, encoding="utf-8-sig",
                     dtype={"codice": str, "piva": str, "dipendente": str})
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def sostituisci(con, tabella, df, chiavi):
    """Cancella le righe con le stesse chiavi (es. stessa data) e inserisce le nuove: niente doppioni."""
    for valori in df[chiavi].drop_duplicates().itertuples(index=False):
        cond = " AND ".join(f"{c}=?" for c in chiavi)
        con.execute(f"DELETE FROM {tabella} WHERE {cond}", tuple(valori))
    df.to_sql(tabella, con, if_exists="append", index=False)


def importa_csv(con, f, tabella):
    df = leggi_csv(f)
    mancanti = set(COLONNE[tabella]) - set(df.columns)
    if mancanti:
        raise ValueError(f"colonne mancanti: {sorted(mancanti)}")
    df = df[COLONNE[tabella]].copy()
    df["data"] = pd.to_datetime(df["data"], dayfirst=True, format="mixed").dt.strftime("%Y-%m-%d")
    sostituisci(con, tabella, df, ["data"])


def importa_inventario(con, f):
    m = re.search(r"(\d{4}-\d{2}-\d{2})", f.name)
    if not m:
        raise ValueError("il nome file deve contenere la data, es. inventario_2026-11-01.csv")
    df = leggi_csv(f)[["ingrediente", "quantita"]].copy()
    df.insert(0, "data", m.group(1))
    sostituisci(con, "inventari", df, ["data"])


def estrai_xml(raw: bytes) -> bytes:
    """Restituisce l'XML della FatturaPA. Gestisce .xml e, nella maggior parte dei casi, .p7m firmati."""
    if raw.lstrip()[:1] == b"<":
        return raw
    m = re.search(rb"<\?xml.*?</[\w:]*FatturaElettronica>", raw, re.S) or \
        re.search(rb"<[\w:]*FatturaElettronica[\s>].*?</[\w:]*FatturaElettronica>", raw, re.S)
    if not m:
        raise ValueError("p7m non leggibile in modo semplice (serve estrazione con openssl)")
    return m.group(0)


def _t(el, percorso, default=""):
    x = el.find(percorso) if el is not None else None
    return x.text.strip() if x is not None and x.text else default


def leggi_fattura(raw: bytes) -> pd.DataFrame:
    """Legge una fattura elettronica (FatturaPA, standard SDI) e restituisce una riga per DettaglioLinee."""
    root = ET.fromstring(estrai_xml(raw))
    for el in root.iter():  # toglie i namespace (p:, ns2:...) per semplificare le ricerche
        el.tag = el.tag.split("}")[-1]
    ced = root.find(".//CedentePrestatore")
    piva = _t(ced, ".//IdFiscaleIVA/IdCodice")
    nome = _t(ced, ".//Denominazione") or f"{_t(ced, './/Nome')} {_t(ced, './/Cognome')}".strip()
    righe = []
    for body in root.iter("FatturaElettronicaBody"):
        doc = body.find("DatiGenerali/DatiGeneraliDocumento")
        tipo = _t(doc, "TipoDocumento")
        segno = -1 if tipo == "TD04" else 1  # TD04 = nota di credito: storna costo e quantità
        for l in body.iter("DettaglioLinee"):
            righe.append(dict(
                piva=piva, fornitore=nome, numero=_t(doc, "Numero"), data=_t(doc, "Data"), tipo=tipo,
                codice_art=_t(l, "CodiceArticolo/CodiceValore"), descrizione=_t(l, "Descrizione"),
                quantita=segno * float(_t(l, "Quantita", "1")), um=_t(l, "UnitaMisura"),
                importo=segno * float(_t(l, "PrezzoTotale", "0")), aliquota=float(_t(l, "AliquotaIVA", "0"))))
    if not righe:
        raise ValueError("nessuna riga di dettaglio trovata")
    return pd.DataFrame(righe)


def importa_fattura(con, f):
    sostituisci(con, "fatture_righe", leggi_fattura(f.read_bytes()), ["piva", "numero", "data"])


# (sottocartella di inbox, schema nome file, funzione che lo importa)
IMPORTATORI = [
    ("cassa", "vendite_*.csv", lambda con, f: importa_csv(con, f, "vendite")),
    ("cassa", "coperti_*.csv", lambda con, f: importa_csv(con, f, "coperti")),
    ("fatture", "*.xml", importa_fattura),
    ("fatture", "*.p7m", importa_fattura),
    ("presenze", "*.csv", lambda con, f: importa_csv(con, f, "presenze")),
    ("inventari", "*.csv", importa_inventario),
]


def importa_tutto(con):
    nuovi, scartati = 0, 0
    for sub, schema, funzione in IMPORTATORI:
        (INBOX / sub).mkdir(exist_ok=True)
        for f in sorted((INBOX / sub).glob(schema)):
            destinazione = ARCHIVIO
            h = hashlib.md5(f.read_bytes()).hexdigest()
            try:
                if con.execute("SELECT 1 FROM file_importati WHERE hash=?", (h,)).fetchone():
                    log.info(f"già importato, salto: {sub}/{f.name}")
                else:
                    funzione(con, f)
                    con.execute("INSERT INTO file_importati VALUES (?,?,?)",
                                (h, f"{sub}/{f.name}", dt.datetime.now().isoformat(timespec="seconds")))
                    con.commit()
                    nuovi += 1
            except Exception as e:  # un file rotto non blocca gli altri
                con.rollback()
                destinazione, scartati = SCARTI, scartati + 1
                log.error(f"SCARTATO {sub}/{f.name}: {e}")
            (destinazione / sub).mkdir(parents=True, exist_ok=True)
            os.replace(f, destinazione / sub / f.name)
    log.info(f"Import: {nuovi} file nuovi, {scartati} scartati")
    return nuovi, scartati


# --------------------------------------------------------------------------
# 3. CALCOLI
# --------------------------------------------------------------------------
def carica(con):
    D = {t: pd.read_sql(f"SELECT * FROM {t}", con)
         for t in ["vendite", "coperti", "fatture_righe", "presenze", "inventari"]}
    for df in D.values():
        df["data"] = pd.to_datetime(df["data"])
    A = {n: leggi_csv(ANAG / f"{n}.csv")
         for n in ["articoli", "ricette", "ingredienti", "mappa_acquisti", "personale", "costi_fissi"]}
    return D, A


def classifica_acquisti(fr, mappa):
    """Assegna a ogni riga fattura ingrediente + categoria di costo usando mappa_acquisti.csv.
    La prima regola che corrisponde vince. 'fattore' converte l'unità del fornitore
    in quella della ricetta (es. sacco 25 kg -> fattore 25)."""
    fr = fr.copy()
    fr["ingrediente"], fr["categoria_costo"], fr["qta_std"] = None, "da_classificare", 0.0
    libero = pd.Series(True, index=fr.index)
    for r in mappa.itertuples():
        m = libero & fr["descrizione"].str.contains(str(r.contiene), case=False, regex=False)
        if str(r.piva) not in ("*", "nan", "", "<NA>"):
            m &= fr["piva"] == str(r.piva)
        fr.loc[m, "categoria_costo"] = r.categoria_costo
        if isinstance(r.ingrediente, str) and r.ingrediente:
            fr.loc[m, "ingrediente"] = r.ingrediente
            fr.loc[m, "qta_std"] = fr.loc[m, "quantita"] * float(r.fattore)
        libero &= ~m
    return fr


def costi_ingredienti(acq, ingr, oggi):
    """Costo unitario = media ponderata degli acquisti recenti; se mancano, il costo manuale."""
    rec = acq[acq.ingrediente.notna() & (acq.qta_std > 0) &
              (acq.data > oggi - pd.Timedelta(days=CFG["giorni_media_costo_acquisto"]))]
    g = rec.groupby("ingrediente").agg(imp=("importo", "sum"), q=("qta_std", "sum"))
    out = ingr.set_index("ingrediente").join((g.imp / g.q).rename("costo_acquisti"))
    out["costo_unitario"] = out.costo_acquisti.fillna(out.costo_unitario_default)
    out["fonte_costo"] = out.costo_acquisti.notna().map({True: "fatture", False: "manuale"})
    return out


def valore_inventari(inv, costi):
    if inv.empty:
        return pd.Series(dtype=float)
    x = inv.merge(costi[["costo_unitario"]], left_on="ingrediente", right_index=True, how="left")
    return (x.quantita * x.costo_unitario).groupby(x.data).sum()


def inventario_vicino(val_inv, giorno, tolleranza=3):
    """Valore dell'inventario fatto entro ±3 giorni da 'giorno' (None se non c'è)."""
    if val_inv.empty:
        return None
    d = (val_inv.index - giorno).days
    ok = val_inv[abs(d) <= tolleranza]
    return float(ok.iloc[0]) if len(ok) else None


def conto_economico(v, acq, pers, val_inv, fissi, ultimo_giorno, n_mesi=6):
    merci = CFG["categorie_costo_merci"]
    fissi_mese = fissi.importo_mensile.sum()
    righe = {}
    for mese in sorted(v.data.dt.to_period("M").unique())[-n_mesi:]:
        inizio, fine = mese.start_time, mese.end_time.normalize()
        in_corso = ultimo_giorno < fine
        quota = (ultimo_giorno.day / mese.days_in_month) if in_corso else 1  # costi fissi pro-rata
        vm = v[v.data.dt.to_period("M") == mese]
        am = acq[acq.data.dt.to_period("M") == mese]
        ricavi = vm.ricavo_netto.sum()
        acq_food = am[am.categoria_costo == merci[0]].importo.sum()
        acq_bev = am[am.categoria_costo.isin(merci[1:])].importo.sum()
        i0, i1 = inventario_vicino(val_inv, inizio), inventario_vicino(val_inv, fine + pd.Timedelta(days=1))
        var_rim = (i0 - i1) if (i0 is not None and i1 is not None) else 0.0
        costo_merci = acq_food + acq_bev + var_rim
        personale = pers[pers.data.dt.to_period("M") == mese].costo.sum()
        altri = am[~am.categoria_costo.isin(merci)].importo.sum()
        fissi_m = fissi_mese * quota
        struttura = personale + altri + fissi_m
        risultato = ricavi - costo_merci - struttura
        mc_pct = 1 - costo_merci / ricavi if ricavi else 0
        righe[str(mese) + (" (in corso)" if in_corso else "")] = {
            "Ricavi netti (IVA esclusa)": ricavi,
            "  di cui non scontrinato": vm[vm.scontrinato.str.upper() == "N"].ricavo_netto.sum(),
            "Coperti": 0,  # riempito dopo dal file coperti
            "Acquisti food": acq_food, "Acquisti beverage": acq_bev,
            "Variazione rimanenze": var_rim,
            "COSTO MERCI": costo_merci,
            "MARGINE LORDO": ricavi - costo_merci,
            "Costo personale": personale,
            "Altri costi da fatture": altri,
            "Costi fissi (anagrafica)": fissi_m,
            "RISULTATO GESTIONALE": risultato,
            "Food cost % effettivo": costo_merci / ricavi if ricavi else 0,
            "Food cost % teorico (ricette)": vm.costo_teorico.sum() / ricavi if ricavi else 0,
            "Costo lavoro %": personale / ricavi if ricavi else 0,
            "Prime cost %": (costo_merci + personale) / ricavi if ricavi else 0,
            "Risultato %": risultato / ricavi if ricavi else 0,
            "Ricavi di pareggio": struttura / mc_pct if mc_pct > 0 else 0,
            "Rimanenze": "con inventari" if (i0 is not None and i1 is not None) else "SENZA inventari (stima)",
        }
    return pd.DataFrame(righe)


def menu_engineering(v, cp, art, oggi):
    """Metodo Kasavana-Smith, separato per gruppo (pizze, birre, cucina...):
    popolarità alta se quota >= 70% della quota media; margine alto se >= margine medio ponderato."""
    d0 = oggi - pd.Timedelta(days=CFG["giorni_menu_engineering"] - 1)
    x = v[(v.data >= d0) & v.codice.isin(cp.index)]
    g = x.groupby("codice").agg(venduti=("quantita", "sum"), ricavo=("ricavo_netto", "sum")).reset_index()
    g = g.merge(art[["codice", "descrizione", "gruppo_me"]], on="codice", how="left")
    g["prezzo_medio_netto"] = g.ricavo / g.venduti
    g["costo_porzione"] = g.codice.map(cp)
    g["food_cost_pct"] = g.costo_porzione / g.prezzo_medio_netto
    g["margine_unitario"] = g.prezzo_medio_netto - g.costo_porzione
    g["margine_totale"] = g.margine_unitario * g.venduti
    out = []
    for _, gr in g.groupby("gruppo_me"):
        gr = gr.copy()
        gr["mix_pct"] = gr.venduti / gr.venduti.sum()
        soglia_pop = CFG.get("soglia_popolarita", 0.7) / len(gr)
        soglia_mc = gr.margine_totale.sum() / gr.venduti.sum()
        alta_pop, alto_mc = gr.mix_pct >= soglia_pop, gr.margine_unitario >= soglia_mc
        gr["classe"] = "Dog"
        gr.loc[alta_pop & alto_mc, "classe"] = "Star"
        gr.loc[alta_pop & ~alto_mc, "classe"] = "Plowhorse"
        gr.loc[~alta_pop & alto_mc, "classe"] = "Puzzle"
        out.append(gr)
    me = pd.concat(out) if out else g
    azioni = {"Star": "Mantieni qualità e posizione in menu",
              "Plowhorse": "Piace ma rende poco: ritocca prezzo o riduci costo ricetta",
              "Puzzle": "Rende ma vende poco: mettila in evidenza / proponila in sala",
              "Dog": "Valuta se toglierla o ripensarla"}
    me["azione"] = me.classe.map(azioni)
    return me.sort_values(["gruppo_me", "margine_totale"], ascending=[True, False])


def consumi_sprechi(v, acq, inv, ricette, costi, oggi):
    """Consumo REALE (inventario iniziale + acquisti - inventario finale) contro consumo TEORICO
    (vendite x ricette). La differenza è spreco/porzionatura/omaggi/ammanchi potenziali."""
    date_inv = sorted(inv.data.unique())
    if len(date_inv) >= 2:
        d0, d1, metodo = date_inv[-2], date_inv[-1], "tra gli ultimi due inventari"
    else:
        d1 = oggi + pd.Timedelta(days=1)
        d0, metodo = d1 - pd.Timedelta(days=30), "ultimi 30 giorni SENZA inventari (solo acquisti: stima)"
    vv = v[(v.data >= d0) & (v.data < d1)]
    teo = vv.merge(ricette, on="codice").assign(q=lambda t: t.quantita_x * t.quantita_y) \
        .groupby("ingrediente").q.sum().rename("teorico")
    acquisti = acq[(acq.data >= d0) & (acq.data < d1) & acq.ingrediente.notna()] \
        .groupby("ingrediente").qta_std.sum().rename("acquistato")
    df = pd.concat([teo, acquisti], axis=1).fillna(0)
    if len(date_inv) >= 2:
        i0 = inv[inv.data == d0].set_index("ingrediente").quantita
        i1 = inv[inv.data == d1].set_index("ingrediente").quantita
        df["inv_iniziale"], df["inv_finale"] = i0.reindex(df.index).fillna(0), i1.reindex(df.index).fillna(0)
    else:
        df["inv_iniziale"] = df["inv_finale"] = 0.0
    df["reale"] = df.inv_iniziale + df.acquistato - df.inv_finale
    df["differenza"] = df.reale - df.teorico
    df["scostamento_pct"] = df.differenza / df.teorico.where(df.teorico > 0)
    df["unita"] = costi.unita.reindex(df.index)
    df["valore_differenza"] = df.differenza * costi.costo_unitario.reindex(df.index)
    df = df.sort_values("valore_differenza", ascending=False)
    periodo = f"{pd.Timestamp(d0):%d/%m/%Y} - {pd.Timestamp(d1) - pd.Timedelta(days=1):%d/%m/%Y} ({metodo})"
    return df, periodo


GIORNI = ["lun", "mar", "mer", "gio", "ven", "sab", "dom"]


def giorno_it(g):
    return f"{GIORNI[g.weekday()]} {g:%d/%m}"


def kpi_finestra(v, cop, pers, d0, d1):
    vv, cc, pp = (x[(x.data >= d0) & (x.data <= d1)] for x in (v, cop, pers))
    ricavi, lordo, coperti, ore = vv.ricavo_netto.sum(), vv.importo_lordo.sum(), cc.coperti.sum(), pp.ore.sum()
    return {
        "Ricavi netti": ricavi, "Coperti": coperti,
        "Scontrino medio/coperto (IVA incl.)": lordo / coperti if coperti else 0,
        "Ore lavorate": ore, "Ricavo per ora lavorata": ricavi / ore if ore else 0,
        "Costo lavoro %": pp.costo.sum() / ricavi if ricavi else 0,
        "Food cost % teorico": vv.costo_teorico.sum() / ricavi if ricavi else 0,
    }


# --------------------------------------------------------------------------
# 4. REPORT (HTML + Excel + nota Obsidian)
# --------------------------------------------------------------------------
def eur(x):
    try:
        return f"€ {x:,.0f}".replace(",", ".") if abs(x) >= 100 else f"€ {x:,.2f}".replace(".", ",")
    except (TypeError, ValueError):
        return str(x)


def pct(x):
    try:
        return f"{x * 100:.1f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "-"


def num(x, dec=0):
    try:
        return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(x)


def tabella(df, formati, classi_riga=None):
    """DataFrame -> tabella HTML con formati per colonna (eur/pct/num)."""
    al = {c: "" if c in formati else ' class="t"' for c in df.columns}  # testo allineato a sinistra
    th = "".join(f"<th{al[c]}>{escape(str(c))}</th>" for c in df.columns)
    righe = []
    for i, (_, r) in enumerate(df.iterrows()):
        celle = "".join(f"<td{al[c]}>{escape(formati.get(c, str)(r[c]) if pd.notna(r[c]) else '-')}</td>" for c in df.columns)
        cls = f' class="{classi_riga[i]}"' if classi_riga else ""
        righe.append(f"<tr{cls}>{celle}</tr>")
    return f"<table><thead><tr>{th}</tr></thead><tbody>{''.join(righe)}</tbody></table>"


# ---- GRAFICI SVG (disegnati dallo script: funzionano offline, nessuna libreria esterna)
import math


def svg_tachimetro(valore, minimo, massimo, zone, titolo, testo, sotto="", w=300):
    """Tachimetro a semicerchio. zone = [(da, a, colore), ...] nello stesso intervallo del valore."""
    cx, cy, r = w / 2, w * 0.55, w * 0.42
    ang = lambda v: math.pi * (1 - (min(max(v, minimo), massimo) - minimo) / (massimo - minimo))
    pt = lambda a, rr: (cx + rr * math.cos(a), cy - rr * math.sin(a))
    archi = ""
    for da, a, col in zone:
        (x1, y1), (x2, y2) = pt(ang(da), r), pt(ang(a), r)
        archi += f'<path d="M{x1:.1f},{y1:.1f} A{r},{r} 0 0 1 {x2:.1f},{y2:.1f}" stroke="{col}" stroke-width="{w*0.07:.0f}" fill="none"/>'
    nx, ny = pt(ang(valore), r * 0.82)
    return (f'<svg viewBox="0 0 {w} {w*0.72:.0f}" class="gauge"><text x="{cx}" y="16" text-anchor="middle" class="gt">{escape(titolo)}</text>'
            f'{archi}<line x1="{cx}" y1="{cy}" x2="{nx:.1f}" y2="{ny:.1f}" stroke="#1d2b36" stroke-width="4" stroke-linecap="round"/>'
            f'<circle cx="{cx}" cy="{cy}" r="7" fill="#1d2b36"/><text x="{cx}" y="{cy+w*0.1:.0f}" text-anchor="middle" class="gv">{escape(testo)}</text>'
            f'<text x="{cx}" y="{cy+w*0.16:.0f}" text-anchor="middle" class="gs">{escape(sotto)}</text></svg>')


def svg_barre(valori, etichette, riferimento=None, rif_testo="", fmt=lambda x: f"{x:.0f}", colori=None, w=640, h=220):
    """Grafico a barre con linea di riferimento opzionale (es. pareggio giornaliero)."""
    if not len(valori):
        return ""
    top = max(max(valori), riferimento or 0) * 1.12 or 1
    bot = min(0, min(valori))
    sx, base = (w - 75) / len(valori), lambda v: 20 + (h - 50) * (top - v) / (top - bot)
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart">']
    for i, (v, e) in enumerate(zip(valori, etichette)):
        x, y0, y1 = 65 + i * sx, base(0), base(v)
        col = (colori[i] if colori else "#c8963e")
        s.append(f'<rect x="{x+sx*0.12:.1f}" y="{min(y0,y1):.1f}" width="{sx*0.76:.1f}" height="{abs(y0-y1):.1f}" fill="{col}"><title>{escape(e)}: {escape(fmt(v))}</title></rect>')
        if len(valori) <= 12 or i % max(1, len(valori) // 10) == 0:
            s.append(f'<text x="{x+sx/2:.1f}" y="{h-8}" text-anchor="middle" class="ax">{escape(e)}</text>')
    s.append(f'<line x1="65" x2="{w-10}" y1="{base(0):.1f}" y2="{base(0):.1f}" stroke="#999" stroke-width=".5"/>')
    s.append(f'<text x="60" y="{base(top/1.12):.0f}" text-anchor="end" class="ax">{escape(fmt(top/1.12))}</text>')
    if riferimento:
        y = base(riferimento)
        s.append(f'<line x1="65" x2="{w-10}" y1="{y:.1f}" y2="{y:.1f}" stroke="#c0392b" stroke-dasharray="5 4" stroke-width="1.5"/>'
                 f'<text x="60" y="{y+4:.1f}" text-anchor="end" class="ax" fill="#c0392b">{escape(rif_testo)}</text>')
    return "".join(s) + "</svg>"


def svg_linee(serie, etichette, soglia=None, w=640, h=200):
    """Linee percentuali (es. food cost % e costo lavoro % per settimana) con soglia obiettivo."""
    tutti = [v for _, vals, _ in serie for v in vals if pd.notna(v)] + ([soglia] if soglia else [])
    if not tutti:
        return ""
    top, n = max(tutti) * 1.15, len(etichette)
    X = lambda i: 45 + i * (w - 125) / max(1, n - 1)
    Y = lambda v: 15 + (h - 45) * (1 - v / top)
    s = [f'<svg viewBox="0 0 {w} {h}" class="chart">']
    if soglia:
        s.append(f'<line x1="45" x2="{X(n-1):.1f}" y1="{Y(soglia):.1f}" y2="{Y(soglia):.1f}" stroke="#c0392b" stroke-dasharray="5 4"/>'
                 f'<text x="40" y="{Y(soglia)+4:.1f}" text-anchor="end" class="ax">{pct(soglia)}</text>')
    for nome, vals, col in serie:
        pts = " ".join(f"{X(i):.1f},{Y(v):.1f}" for i, v in enumerate(vals) if pd.notna(v))
        s.append(f'<polyline points="{pts}" fill="none" stroke="{col}" stroke-width="2.5"/>')
        s.append(f'<text x="{X(n-1)+8:.1f}" y="{Y(vals[-1])+4:.1f}" class="ax" fill="{col}">{escape(nome)}</text>')
    for i, e in enumerate(etichette):
        if n <= 12 or i % 2 == 0:
            s.append(f'<text x="{X(i):.1f}" y="{h-8}" text-anchor="middle" class="ax">{escape(e)}</text>')
    return "".join(s) + "</svg>"


CSS = """
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f5f3ef;color:#222}
header{background:#1d2b36;color:#fff;padding:18px 28px}header h1{margin:0;font-size:22px}
header p{margin:4px 0 0;opacity:.75;font-size:13px}main{padding:20px 28px;max-width:1300px}
h2{font-size:17px;border-bottom:2px solid #c8963e;padding-bottom:4px;margin-top:34px}
.cards{display:flex;flex-wrap:wrap;gap:12px}.card{background:#fff;border-radius:8px;padding:12px 16px;
min-width:170px;box-shadow:0 1px 3px #0001}.card b{display:block;font-size:20px}.card span{font-size:12px;color:#666}
.card.ko{border-left:5px solid #c0392b}.card.ok{border-left:5px solid #27ae60}
table{border-collapse:collapse;background:#fff;font-size:13px;margin:8px 0;box-shadow:0 1px 3px #0001}
th,td{padding:5px 9px;border-bottom:1px solid #eee;text-align:right}th{background:#1d2b36;color:#fff;position:sticky;top:0}
td:first-child,th:first-child,td.t,th.t{text-align:left}tr.tot td{font-weight:bold;background:#f0ebe1}
tr.ko td{background:#fdecea}.avvisi li{margin:4px 0}.avvisi .alto{color:#c0392b;font-weight:bold}
.matrice{display:grid;grid-template-columns:1fr 1fr;gap:8px;max-width:700px}.q{background:#fff;padding:10px;border-radius:8px}
.q h3{margin:0 0 6px;font-size:14px}.Star{border-top:4px solid #27ae60}.Plowhorse{border-top:4px solid #2980b9}
.Puzzle{border-top:4px solid #e67e22}.Dog{border-top:4px solid #c0392b}small{color:#666}
.gauges{display:flex;flex-wrap:wrap;gap:12px}.gauges>div{background:#fff;border-radius:8px;padding:8px;box-shadow:0 1px 3px #0001}
.gauge{width:300px;max-width:100%}.gauge.s{width:220px}.gt{font:600 14px Segoe UI,Arial}.gv{font:700 22px Segoe UI,Arial}
.gs{font:12px Segoe UI,Arial;fill:#666}.chart{width:100%;max-width:820px;background:#fff;border-radius:8px;box-shadow:0 1px 3px #0001}
.ax{font:11px Segoe UI,Arial;fill:#666}.chart{display:block}
table{display:block;overflow-x:auto;max-width:100%}@media(max-width:600px){main{padding:12px}.gauge.s{width:46%}}
"""


def card(titolo, valore, stato=""):
    return f'<div class="card {stato}"><span>{escape(titolo)}</span><b>{escape(valore)}</b></div>'


def stato(val, soglia):
    return "ko" if val > soglia else "ok"


def scrivi_report(R):
    oggi = R["ultimo_giorno"]
    k = R["kpi"]
    h = [f"<!doctype html><html lang='it'><head><meta charset='utf-8'><title>Galleon - Report</title>"
         f"<meta name='viewport' content='width=device-width,initial-scale=1'><style>{CSS}</style></head><body>"
         f"<header><h1>The Galleon - Report gestionale</h1><p>Dati fino al {oggi:%d/%m/%Y} · "
         f"generato il {dt.datetime.now():%d/%m/%Y %H:%M} · tutto calcolato in locale</p></header><main>"]

    # Avvisi
    h.append("<h2>Cosa guardare oggi</h2><ul class='avvisi'>")
    h += [f"<li class='{liv}'>{escape(t)}</li>" for liv, t in R["avvisi"]] or ["<li>Nessuna criticità.</li>"]
    h.append("</ul>")

    # Bussola: dove sto andando questo mese
    b = R["bussola"]
    VERDE, GIALLO, ROSSO = "#27ae60", "#f1c40f", "#c0392b"
    h.append("<h2>Bussola del mese (proiezione a fine mese)</h2><div class='gauges'>")
    h.append("<div>" + svg_tachimetro(b["rapporto"] * 100, 50, 200, [(50, 100, ROSSO), (100, 120, GIALLO), (120, 200, VERDE)],
             "Rotta verso l'utile", f"{b['rapporto']*100:.0f}% del pareggio",
             f"Risultato stimato {eur(b['risultato'])}") + "</div>")
    fc, cl = OBIETTIVI["food_cost_max"] * 100, OBIETTIVI["costo_lavoro_max"] * 100
    for tit, val, sog in (("Food cost", b["food_cost"], fc), ("Costo lavoro", b["costo_lavoro"], cl)):
        h.append("<div>" + svg_tachimetro(val * 100, 10, 50, [(10, sog - 3, VERDE), (sog - 3, sog, GIALLO), (sog, 50, ROSSO)],
                 tit, pct(val), f"obiettivo max {sog:.0f}%").replace('class="gauge"', 'class="gauge s"') + "</div>")
    h.append("</div><small>Rotta = ricavi proiettati a fine mese ÷ ricavi di pareggio. Sotto 100% = perdita, "
             "100-120% = margine stretto, sopra 120% = in rotta.</small>")

    h.append("<h2>Incassi giornalieri (ultimi 30 giorni)</h2>")
    gg = R["serie_giorni"]
    h.append(svg_barre(list(gg.ricavi), [f"{d:%d/%m}" for d in gg.index], b["pareggio_giorno"],
                       "pareggio", eur,
                       ["#27ae60" if v >= b["pareggio_giorno"] else "#c8963e" for v in gg.ricavi]))
    h.append(f"<small>Linea rossa = incasso netto medio per serata che serve per andare in pari "
             f"(≈ {eur(b['pareggio_giorno'])}). Verde = serata sopra il pareggio.</small>")
    h.append("<h2>Food cost e costo lavoro per settimana</h2>")
    sw = R["serie_settimane"]
    h.append(svg_linee([("Food cost", list(sw.food_cost), "#c8963e"), ("Lavoro", list(sw.lavoro), "#2980b9")],
                       [f"{d - pd.Timedelta(days=6):%d/%m}" for d in sw.index], OBIETTIVI["food_cost_max"]))
    h.append("<small>Ogni punto è una settimana (data di inizio). Linea rossa = obiettivo massimo.</small>")
    h.append("<h2>Risultato gestionale per mese</h2>")
    ris = R["ce"].loc["RISULTATO GESTIONALE"]
    h.append(svg_barre(list(ris.values), list(ris.index), fmt=eur,
                       colori=["#27ae60" if v >= 0 else "#c0392b" for v in ris.values], h=200))

    # KPI
    for nome, kk in k.items():
        h.append(f"<h2>{escape(nome)}</h2><div class='cards'>")
        h.append(card("Ricavi netti", eur(kk["Ricavi netti"])))
        h.append(card("Coperti", num(kk["Coperti"])))
        h.append(card("Scontrino medio/coperto", eur(kk["Scontrino medio/coperto (IVA incl.)"])))
        h.append(card("Ricavo per ora lavorata", eur(kk["Ricavo per ora lavorata"])))
        h.append(card("Costo lavoro %", pct(kk["Costo lavoro %"]), stato(kk["Costo lavoro %"], OBIETTIVI["costo_lavoro_max"])))
        h.append(card("Food cost % teorico", pct(kk["Food cost % teorico"]), stato(kk["Food cost % teorico"], OBIETTIVI["food_cost_max"])))
        h.append("</div>")

    # Conto economico
    ce = R["ce"]
    h.append("<h2>Conto economico gestionale per mese</h2>")
    ce_html = ce.copy().astype(object)
    for idx in ce.index:
        if "%" in idx:
            ce_html.loc[idx] = [pct(x) for x in ce.loc[idx]]
        elif idx == "Coperti":
            ce_html.loc[idx] = [num(x) for x in ce.loc[idx]]
        elif idx != "Rimanenze":
            ce_html.loc[idx] = [eur(x) for x in ce.loc[idx]]
    ce_html = ce_html.reset_index().rename(columns={"index": "Voce"})
    cls = ["tot" if v.isupper() or v.startswith(("MARGINE", "RISULTATO", "COSTO")) else "" for v in ce_html.Voce]
    h.append(tabella(ce_html, {c: str for c in ce_html.columns[1:]}, cls))
    h.append("<small>Ricavi di pareggio = ricavi netti minimi del mese per chiudere a zero con questi costi. "
             "Il mese in corso ha i costi fissi calcolati pro-rata.</small>")

    # Andamento giornaliero
    h.append("<h2>Ultimi 14 giorni</h2>")
    h.append(tabella(R["giorni"], {"Ricavi netti": eur, "Coperti": num, "Ore": lambda x: num(x, 1),
                                   "Costo lavoro %": pct, "Food cost teorico %": pct}))

    # Menu engineering
    me = R["me"]
    h.append(f"<h2>Menu engineering (ultimi {CFG['giorni_menu_engineering']} giorni, prezzi e costi aggiornati)</h2>")
    h.append("<div class='matrice'>")
    for cl, tit in [("Puzzle", "Puzzle - margine alto, vendite basse"), ("Star", "Star - margine alto, vendite alte"),
                    ("Dog", "Dog - margine basso, vendite basse"), ("Plowhorse", "Plowhorse - margine basso, vendite alte")]:
        nomi = ", ".join(me[me.classe == cl].descrizione) or "-"
        h.append(f"<div class='q {cl}'><h3>{tit}</h3>{escape(nomi)}</div>")
    h.append("</div>")
    cols = ["gruppo_me", "descrizione", "venduti", "mix_pct", "prezzo_medio_netto", "costo_porzione",
            "food_cost_pct", "margine_unitario", "margine_totale", "classe", "azione"]
    t = me[cols].rename(columns={"gruppo_me": "Gruppo", "descrizione": "Piatto", "venduti": "Venduti", "mix_pct": "Mix %",
                                 "prezzo_medio_netto": "Prezzo netto", "costo_porzione": "Costo porzione",
                                 "food_cost_pct": "Food cost %", "margine_unitario": "Margine unit.",
                                 "margine_totale": "Margine tot.", "classe": "Classe", "azione": "Azione"})
    cls = ["ko" if x > OBIETTIVI["food_cost_max"] + 0.05 else "" for x in t["Food cost %"]]
    h.append(tabella(t, {"Venduti": num, "Mix %": pct, "Prezzo netto": eur, "Costo porzione": eur,
                         "Food cost %": pct, "Margine unit.": eur, "Margine tot.": eur}, cls))

    # Consumi
    cs, periodo = R["consumi"]
    h.append(f"<h2>Controllo consumi e sprechi</h2><small>Periodo: {escape(periodo)}. "
             "Positivo = consumato più di quanto le vendite giustificano.</small>")
    t = cs.reset_index().rename(columns={"index": "ingrediente"})
    t = t[["ingrediente", "unita", "teorico", "inv_iniziale", "acquistato", "inv_finale", "reale", "differenza",
           "scostamento_pct", "valore_differenza"]]
    cls = ["ko" if (pd.notna(s) and s > OBIETTIVI["scostamento_consumi_max"] and v > 20) else ""
           for s, v in zip(t.scostamento_pct, t.valore_differenza)]
    f1 = lambda x: num(x, 1)
    h.append(tabella(t, {"teorico": f1, "inv_iniziale": f1, "acquistato": f1, "inv_finale": f1, "reale": f1,
                         "differenza": f1, "scostamento_pct": pct, "valore_differenza": eur}, cls))

    # Costi ingredienti
    h.append("<h2>Costo ingredienti in uso</h2>")
    ci = R["costi"].reset_index()[["ingrediente", "unita", "costo_unitario", "costo_unitario_default", "fonte_costo"]]
    h.append(tabella(ci, {"costo_unitario": eur, "costo_unitario_default": eur}))

    # Qualità dati
    h.append("<h2>Qualità dei dati (da sistemare per avere numeri affidabili)</h2><ul>")
    h += [f"<li>{escape(x)}</li>" for x in R["qualita"]] or ["<li>Tutto in ordine.</li>"]
    h.append("</ul></main></body></html>")

    html = "".join(h)
    (REPORT / "Galleon_report.html").write_text(html, encoding="utf-8")
    storico = REPORT / "storico"
    storico.mkdir(exist_ok=True)
    (storico / f"Galleon_report_{oggi:%Y-%m-%d}.html").write_text(html, encoding="utf-8")

    # Excel: stessi dati, per chi vuole filtrare/ordinare
    with pd.ExcelWriter(REPORT / "Galleon_report.xlsx", engine="openpyxl") as xw:
        R["ce"].to_excel(xw, sheet_name="Conto economico")
        R["giorni"].to_excel(xw, sheet_name="Ultimi 14 giorni", index=False)
        R["me"][cols].to_excel(xw, sheet_name="Menu engineering", index=False)
        R["consumi"][0].to_excel(xw, sheet_name="Consumi e sprechi")
        R["costi"].to_excel(xw, sheet_name="Costi ingredienti")
        R["non_classificati"].to_excel(xw, sheet_name="Fatture da classificare", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(45, max(10, *(len(str(c.value or "")) for c in col)))

    # Nota Obsidian (cruscotto sempre aggiornato, con link al report completo)
    if CFG.get("cartella_obsidian"):
        vault = Path(CFG["cartella_obsidian"])
        vault.mkdir(parents=True, exist_ok=True)
        kk = k["Mese in corso"]
        md = [f"# Galleon - Cruscotto\n_Aggiornato: {dt.datetime.now():%d/%m/%Y %H:%M} - dati fino al {oggi:%d/%m/%Y}_\n",
              f"[Apri report completo](file:///{(REPORT / 'Galleon_report.html').as_posix()})\n",
              "## Mese in corso",
              f"- Ricavi netti: {eur(kk['Ricavi netti'])}", f"- Coperti: {num(kk['Coperti'])}",
              f"- Costo lavoro: {pct(kk['Costo lavoro %'])}", f"- Food cost teorico: {pct(kk['Food cost % teorico'])}",
              "\n## Avvisi"] + [f"- {t}" for _, t in R["avvisi"]]
        (vault / "Galleon - Cruscotto.md").write_text("\n".join(md), encoding="utf-8")


def esporta_app(R):
    """Scrive report/app/dati.js (i numeri già calcolati) + copia l'app grafica app/index.html.
    L'app si apre con doppio clic, anche offline: nessun dato grezzo, solo sintesi."""
    f = lambda x: None if pd.isna(x) else round(float(x), 4)
    ce = R["ce"]
    mesi = [dict(mese=c, in_corso="in corso" in c, ricavi=f(ce.loc["Ricavi netti (IVA esclusa)", c]),
                 non_scontrinato=f(ce.loc["  di cui non scontrinato", c]), coperti=f(ce.loc["Coperti", c]),
                 costo_merci=f(ce.loc["COSTO MERCI", c]), personale=f(ce.loc["Costo personale", c]),
                 altri=f(ce.loc["Altri costi da fatture", c]), fissi=f(ce.loc["Costi fissi (anagrafica)", c]),
                 risultato=f(ce.loc["RISULTATO GESTIONALE", c]), food_eff=f(ce.loc["Food cost % effettivo", c]),
                 food_teo=f(ce.loc["Food cost % teorico (ricette)", c]), lavoro=f(ce.loc["Costo lavoro %", c]),
                 prime=f(ce.loc["Prime cost %", c]), pareggio=f(ce.loc["Ricavi di pareggio", c]),
                 rimanenze=ce.loc["Rimanenze", c]) for c in ce.columns]
    me = R["me"]
    cs, periodo = R["consumi"]
    dati = dict(
        generato=f"{dt.datetime.now():%d/%m/%Y %H:%M}", ultimo_giorno=f"{R['ultimo_giorno']:%Y-%m-%d}",
        obiettivi=OBIETTIVI, avvisi=[dict(alto=a == "alto", testo=t) for a, t in R["avvisi"]],
        bussola={k: f(v) for k, v in R["bussola"].items()},
        kpi={k: {kk: f(vv) for kk, vv in v.items()} for k, v in R["kpi"].items()},
        giorni=[dict(data=f"{d:%Y-%m-%d}", ricavi=f(r)) for d, r in R["serie_giorni"].ricavi.items()],
        settimane=[dict(da=f"{d - pd.Timedelta(days=6):%Y-%m-%d}", food=f(r.food_cost), lavoro=f(r.lavoro))
                   for d, r in R["serie_settimane"].iterrows()],
        mesi=mesi,
        menu=[dict(gruppo=r.gruppo_me, nome=r.descrizione, venduti=f(r.venduti), mix=f(r.mix_pct),
                   prezzo=f(r.prezzo_medio_netto), costo=f(r.costo_porzione), food=f(r.food_cost_pct),
                   margine=f(r.margine_unitario), margine_tot=f(r.margine_totale), classe=r.classe, azione=r.azione)
              for r in me.itertuples()],
        consumi=dict(periodo=periodo, righe=[dict(ingrediente=i, unita=r.unita, teorico=f(r.teorico), reale=f(r.reale),
                     diff=f(r.differenza), pct=f(r.scostamento_pct), valore=f(r.valore_differenza)) for i, r in cs.iterrows()]),
        qualita=R["qualita"])
    out = REPORT / "app"
    out.mkdir(exist_ok=True)
    (out / "dati.js").write_text("window.GALLEON = " + json.dumps(dati, ensure_ascii=False) + ";", encoding="utf-8")
    sorgente = BASE / "app" / "index.html"
    if sorgente.exists():
        shutil.copy(sorgente, out / "index.html")


# --------------------------------------------------------------------------
# 5. ORCHESTRAZIONE
# --------------------------------------------------------------------------
def analizza(con):
    D, A = carica(con)
    v, cop, pres, inv = D["vendite"], D["coperti"], D["presenze"], D["inventari"]
    if v.empty:
        raise SystemExit("Nessuna vendita nel database: metti i file in dati/inbox/cassa e rilancia.")
    oggi = v.data.max()  # ultimo giorno con chiusura disponibile
    art, ric = A["articoli"], A["ricette"]

    acq = classifica_acquisti(D["fatture_righe"], A["mappa_acquisti"])
    costi = costi_ingredienti(acq, A["ingredienti"], oggi)
    r = ric.merge(costi[["costo_unitario"]], left_on="ingrediente", right_index=True, how="left")
    cp = (r.quantita * r.costo_unitario).groupby(r.codice).sum()  # costo porzione per articolo

    if not CFG["includi_non_scontrinato_nei_ricavi"]:
        v = v[v.scontrinato.str.upper() != "N"]
    v = v.merge(art[["codice", "aliquota_iva"]], on="codice", how="left")
    v["aliquota_iva"] = v.aliquota_iva.fillna(CFG["aliquota_iva_default"])
    v["ricavo_netto"] = v.importo_lordo / (1 + v.aliquota_iva / 100)
    v["costo_teorico"] = v.quantita * v.codice.map(cp)

    pers = pres.merge(A["personale"], on="dipendente", how="left")
    pers["costo"] = pers.ore * pers.costo_orario

    val_inv = valore_inventari(inv, costi)
    ce = conto_economico(v, acq, pers, val_inv, A["costi_fissi"], oggi)
    # coperti per mese (aggiunti a parte perché vengono da un altro file)
    cop_m = cop.groupby(cop.data.dt.to_period("M")).coperti.sum()
    ce.loc["Coperti"] = [cop_m.get(pd.Period(c.split(" ")[0], "M"), 0) for c in ce.columns]

    inizio_mese = oggi.replace(day=1)
    kpi = {
        f"Ultima serata ({giorno_it(oggi)})": kpi_finestra(v, cop, pers, oggi, oggi),
        "Ultimi 7 giorni": kpi_finestra(v, cop, pers, oggi - pd.Timedelta(days=6), oggi),
        "Mese in corso": kpi_finestra(v, cop, pers, inizio_mese, oggi),
    }
    giorni = []
    for g in pd.date_range(oggi - pd.Timedelta(days=13), oggi)[::-1]:
        kk = kpi_finestra(v, cop, pers, g, g)
        if kk["Ricavi netti"] or kk["Ore lavorate"]:
            giorni.append({"Giorno": giorno_it(g), "Ricavi netti": kk["Ricavi netti"], "Coperti": kk["Coperti"],
                           "Ore": kk["Ore lavorate"], "Costo lavoro %": kk["Costo lavoro %"],
                           "Food cost teorico %": kk["Food cost % teorico"]})
    # ---- bussola: proiezione del mese in corso
    col = ce.columns[-1]
    fattore = oggi.days_in_month / oggi.day if "in corso" in col else 1
    pareggio_mese = ce.loc["Ricavi di pareggio", col] * fattore
    ricavi_proj = ce.loc["Ricavi netti (IVA esclusa)", col] * fattore
    aperti = v[v.data >= inizio_mese].data.nunique() or 1
    bussola = dict(rapporto=ricavi_proj / pareggio_mese if pareggio_mese else 0,
                   risultato=ce.loc["RISULTATO GESTIONALE", col] * fattore,
                   food_cost=ce.loc["Food cost % teorico (ricette)", col], costo_lavoro=ce.loc["Costo lavoro %", col],
                   pareggio_giorno=ce.loc["Ricavi di pareggio", col] / aperti)
    # ---- serie per i grafici
    d30 = oggi - pd.Timedelta(days=29)
    serie_giorni = v[v.data >= d30].groupby("data").ricavo_netto.sum().rename("ricavi").to_frame()
    sett = v.groupby(pd.Grouper(key="data", freq="W-SUN")).agg(r=("ricavo_netto", "sum"), c=("costo_teorico", "sum"))
    sett["p"] = pers.groupby(pd.Grouper(key="data", freq="W-SUN")).costo.sum()
    sett = sett[sett.r > 0].tail(12)
    serie_settimane = pd.DataFrame({"food_cost": sett.c / sett.r, "lavoro": sett.p / sett.r}, index=sett.index)

    me = menu_engineering(v, cp, art, oggi)
    consumi = consumi_sprechi(v, acq, inv, ric, costi, oggi)

    # ---- avvisi (ordinati: prima i più gravi)
    avvisi = []
    km = kpi["Mese in corso"]
    if km["Costo lavoro %"] > OBIETTIVI["costo_lavoro_max"]:
        avvisi.append(("alto", f"Costo lavoro del mese al {pct(km['Costo lavoro %'])} (obiettivo max {pct(OBIETTIVI['costo_lavoro_max'])})."))
    if km["Food cost % teorico"] > OBIETTIVI["food_cost_max"]:
        avvisi.append(("alto", f"Food cost teorico del mese al {pct(km['Food cost % teorico'])} (obiettivo max {pct(OBIETTIVI['food_cost_max'])})."))
    chiuso = [c for c in ce.columns if "in corso" not in c]
    if chiuso:
        u = ce[chiuso[-1]]
        segno = "in UTILE" if u["RISULTATO GESTIONALE"] >= 0 else "in PERDITA"
        avvisi.append(("alto" if segno == "in PERDITA" else "", f"Ultimo mese chiuso ({chiuso[-1]}): {segno} di {eur(u['RISULTATO GESTIONALE'])}, "
                       f"food cost effettivo {pct(u['Food cost % effettivo'])} vs teorico {pct(u['Food cost % teorico (ricette)'])}."))
    cs = consumi[0]
    for ing, rr in cs[(cs.scostamento_pct > OBIETTIVI["scostamento_consumi_max"]) & (cs.valore_differenza > 20)].head(5).iterrows():
        avvisi.append(("alto", f"{ing}: consumato {pct(rr.scostamento_pct)} in più del teorico (≈ {eur(rr.valore_differenza)} nel periodo)."))
    for _, rr in me[me.food_cost_pct > OBIETTIVI["food_cost_max"] + 0.05].iterrows():
        avvisi.append(("", f"{rr.descrizione}: food cost {pct(rr.food_cost_pct)} - valuta prezzo/ricetta."))
    for _, rr in costi[(costi.fonte_costo == "fatture") &
                       (costi.costo_unitario > costi.costo_unitario_default * 1.05)].iterrows():
        avvisi.append(("", f"{rr.name}: costo d'acquisto salito a {eur(rr.costo_unitario)}/{rr.unita} "
                           f"(+{pct(rr.costo_unitario / rr.costo_unitario_default - 1)} sul costo di riferimento)."))

    # ---- qualità dati
    qualita = []
    nc = acq[acq.categoria_costo == "da_classificare"]
    if len(nc):
        qualita.append(f"{len(nc)} righe fattura non classificate ({eur(nc.importo.sum())}): aggiungi una regola in "
                       f"anagrafiche/mappa_acquisti.csv (elenco nel foglio Excel 'Fatture da classificare').")
    senza_ric = sorted(set(v.codice) - set(ric.codice))
    if senza_ric:
        qualita.append(f"Articoli venduti senza ricetta (food cost non calcolabile): {', '.join(senza_ric)}")
    senza_art = sorted(set(v.codice) - set(art.codice))
    if senza_art:
        qualita.append(f"Codici cassa non presenti in anagrafiche/articoli.csv: {', '.join(senza_art)}")
    no_costo = sorted(pers[pers.costo_orario.isna()].dipendente.unique())
    if no_costo:
        qualita.append(f"Dipendenti senza costo orario in personale.csv: {', '.join(no_costo)}")
    giorni_senza_presenze = sorted(set(v.data.dt.date) - set(pres.data.dt.date))
    if giorni_senza_presenze:
        qualita.append(f"{len(giorni_senza_presenze)} serate con vendite ma senza presenze (ultima: {giorni_senza_presenze[-1]:%d/%m/%Y}).")
    if inv.empty:
        qualita.append("Nessun inventario caricato: food cost effettivo e sprechi sono stime basate solo sugli acquisti.")
    elif (oggi - inv.data.max()).days > 35:
        qualita.append(f"Ultimo inventario del {inv.data.max():%d/%m/%Y}: fanne uno a inizio mese.")
    scarti = list(SCARTI.rglob("*.*"))
    if scarti:
        qualita.append(f"{len(scarti)} file in dati/scarti non importati (motivo in log/pipeline.log).")

    return dict(ultimo_giorno=oggi, kpi=kpi, bussola=bussola, serie_giorni=serie_giorni,
                serie_settimane=serie_settimane, ce=ce, giorni=pd.DataFrame(giorni), me=me, consumi=consumi,
                costi=costi, avvisi=sorted(avvisi, key=lambda a: a[0] != "alto"), qualita=qualita,
                non_classificati=nc[["data", "fornitore", "piva", "numero", "descrizione", "quantita", "importo"]])


def main():
    log.info("=== Avvio pipeline ===")
    try:
        scarica_allegati_email()
    except Exception as e:
        log.error(f"Download email non riuscito (proseguo con i file già presenti): {e}")
    con = sqlite3.connect(DB)
    con.executescript(SCHEMA)
    importa_tutto(con)
    R = analizza(con)
    scrivi_report(R)
    esporta_app(R)
    con.close()
    log.info(f"Report aggiornato: {REPORT / 'Galleon_report.html'}")


if __name__ == "__main__":
    main()
