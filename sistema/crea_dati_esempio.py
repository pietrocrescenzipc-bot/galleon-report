# -*- coding: utf-8 -*-
"""
Crea DATI DI ESEMPIO realistici per provare il sistema (settembre -> ieri):
  - anagrafiche/  : articoli, ricette, ingredienti, mappa acquisti, personale, costi fissi
  - dati/inbox/   : chiusure di cassa giornaliere, fatture elettroniche XML (FatturaPA),
                    presenze mensili, inventari di inizio mese
Nei dati è nascosto di proposito: uno spreco di mozzarella (+12%), un rincaro della
mozzarella a novembre, un piatto poco venduto e righe fattura non classificate.

Uso:  python crea_dati_esempio.py      (ATTENZIONE: sovrascrive anagrafiche e azzera il database)
"""
import math
import random
import shutil
from datetime import date, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
random.seed(7)
INIZIO, FINE = date(2026, 9, 1), date(2026, 11, 18)

# codice: (descrizione, categoria cassa, gruppo menu engineering, prezzo, vendite medie/serata)
ARTICOLI = {
    "P01": ("Margherita", "Pizze", "Pizze", 7.0, 30), "P02": ("Diavola", "Pizze", "Pizze", 8.5, 24),
    "P03": ("Capricciosa", "Pizze", "Pizze", 9.5, 14), "P04": ("Quattro formaggi", "Pizze", "Pizze", 9.0, 9),
    "P05": ("Pizza Galleon al tartufo", "Pizze", "Pizze", 13.0, 5), "P06": ("Bufalina", "Pizze", "Pizze", 10.0, 12),
    "P07": ("Ortolana", "Pizze", "Pizze", 9.0, 2),
    "B01": ("Birra spina media", "Birre", "Bevande", 5.5, 40), "B02": ("Birra spina piccola", "Birre", "Bevande", 3.5, 15),
    "B03": ("Birra artigianale 33cl", "Birre", "Bevande", 7.0, 9), "B04": ("Coca Cola vetro", "Bibite", "Bevande", 3.0, 20),
    "C01": ("Hamburger Galleon", "Cucina", "Cucina", 12.0, 12), "C02": ("Patatine fritte", "Cucina", "Cucina", 4.5, 25),
    "C03": ("Fish & chips", "Cucina", "Cucina", 14.0, 3), "C04": ("Insalatona", "Cucina", "Cucina", 9.0, 4),
}
# ingrediente: (unità, costo di riferimento, fornitore, descrizione in fattura, unità fornitore, fattore conversione)
# Analisi ABC: A = conta settimanale, B = mensile, C = non si conta
CLASSE = {"farina": "A", "mozzarella": "A", "pomodoro": "A", "carne_burger": "A", "bufala": "A", "patatine": "A",
          "salame_piccante": "B", "prosciutto_cotto": "B", "gorgonzola": "B", "fontina": "B", "parmigiano": "B",
          "crema_tartufo": "B", "verdure_grigliate": "B", "olio": "B", "merluzzo": "B", "pane_burger": "B",
          "birra_spina": "B", "birra_bottiglia": "B", "coca": "B"}

ING = {
    "farina": ("kg", 0.90, "A", "FARINA TIPO 00 SACCO 25KG", "SC", 25),
    "mozzarella": ("kg", 6.50, "B", "MOZZARELLA FIORDILATTE FILONE", "KG", 1),
    "pomodoro": ("kg", 1.80, "A", "POLPA POMODORO LATTA 4,1KG", "PZ", 4.1),
    "salame_piccante": ("kg", 14.0, "B", "SALAME PICCANTE SPIANATA", "KG", 1),
    "prosciutto_cotto": ("kg", 11.0, "B", "PROSCIUTTO COTTO ALTA QUALITA", "KG", 1),
    "funghi": ("kg", 5.0, "A", "FUNGHI CHAMPIGNON TRIFOLATI", "KG", 1),
    "carciofi": ("kg", 8.0, "A", "CARCIOFINI SPICCHI OLIO", "KG", 1),
    "olive": ("kg", 6.0, "A", "OLIVE NERE DENOCCIOLATE", "KG", 1),
    "gorgonzola": ("kg", 9.0, "B", "GORGONZOLA DOLCE DOP", "KG", 1),
    "fontina": ("kg", 10.0, "B", "FONTINA DOP", "KG", 1),
    "parmigiano": ("kg", 16.0, "B", "PARMIGIANO REGGIANO GRATTUGIATO", "KG", 1),
    "crema_tartufo": ("kg", 60.0, "A", "CREMA AL TARTUFO 500G", "PZ", 0.5),
    "bufala": ("kg", 12.0, "B", "MOZZARELLA DI BUFALA CAMPANA DOP", "KG", 1),
    "verdure_grigliate": ("kg", 7.0, "A", "VERDURE GRIGLIATE MISTE", "KG", 1),
    "olio": ("l", 9.0, "A", "OLIO EXTRAVERGINE OLIVA 5L", "PZ", 5),
    "birra_spina": ("l", 2.6, "C", "FUSTO BIRRA LAGER 24L", "PZ", 24),
    "birra_bottiglia": ("pz", 2.2, "C", "BIRRA ARTIGIANALE 33CL CT12", "CT", 12),
    "coca": ("pz", 0.7, "C", "COCA COLA VETRO 33CL CT24", "CT", 24),
    "pane_burger": ("pz", 0.5, "A", "PANE BURGER SESAMO", "PZ", 1),
    "carne_burger": ("kg", 13.0, "B", "HAMBURGER SCOTTONA 200G", "KG", 1),
    "patatine": ("kg", 2.4, "A", "PATATE STICK SURGELATE 2,5KG", "PZ", 2.5),
    "merluzzo": ("kg", 15.0, "B", "FILETTI MERLUZZO SURGELATI", "KG", 1),
    "insalata": ("kg", 4.0, "A", "INSALATA MISTA IV GAMMA", "KG", 1),
}
BASE_PIZZA = {"farina": .17, "pomodoro": .08, "mozzarella": .12, "olio": .01}
RICETTE = {
    "P01": BASE_PIZZA, "P02": {**BASE_PIZZA, "salame_piccante": .05},
    "P03": {**BASE_PIZZA, "prosciutto_cotto": .04, "funghi": .03, "carciofi": .03, "olive": .015},
    "P04": {"farina": .17, "mozzarella": .10, "gorgonzola": .04, "fontina": .04, "parmigiano": .02, "olio": .01},
    "P05": {**BASE_PIZZA, "crema_tartufo": .03, "funghi": .04},
    "P06": {"farina": .17, "pomodoro": .08, "bufala": .13, "olio": .015},
    "P07": {**BASE_PIZZA, "verdure_grigliate": .12},
    "B01": {"birra_spina": .4}, "B02": {"birra_spina": .2}, "B03": {"birra_bottiglia": 1}, "B04": {"coca": 1},
    "C01": {"pane_burger": 1, "carne_burger": .2, "patatine": .12, "insalata": .02},
    "C02": {"patatine": .2, "olio": .02}, "C03": {"merluzzo": .2, "patatine": .15, "farina": .03, "olio": .03},
    "C04": {"insalata": .15, "verdure_grigliate": .05, "mozzarella": .05, "olio": .015},
}
FORNITORI = {"A": ("01234560541", "Alimentari Umbria Srl"), "B": ("02345670543", "Caseificio e Salumi Tevere Snc"),
             "C": ("03456780549", "Beverage Centro Italia Spa")}
PERSONALE = {"Marco (pizzaiolo)": (18.0, "pizzeria", 8, "tutti"), "Luca (pizzaiolo)": (15.0, "pizzeria", 6, "weekend"),
             "Sara (sala)": (14.0, "sala", 7, "tutti"), "Giulia (sala)": (13.0, "sala", 6, "weekend"),
             "Ahmed (cucina)": (15.0, "cucina", 7, "tutti")}
SPRECO = {"mozzarella": 1.12}  # si consuma il 12% in più del teorico


def csv(path, intestazione, righe):
    path.parent.mkdir(parents=True, exist_ok=True)
    fmt = lambda x: f"{x:.3f}".rstrip("0").rstrip(".").replace(".", ",") if isinstance(x, float) else str(x)
    path.write_text("\n".join([";".join(intestazione)] + [";".join(fmt(c) for c in r) for r in righe]), encoding="utf-8")


def fattura_xml(forn, numero, giorno, righe):
    piva, nome = FORNITORI[forn]
    linee = "".join(
        f"<DettaglioLinee><NumeroLinea>{i}</NumeroLinea><Descrizione>{d}</Descrizione><Quantita>{q:.2f}</Quantita>"
        f"<UnitaMisura>{um}</UnitaMisura><PrezzoUnitario>{p:.4f}</PrezzoUnitario><PrezzoTotale>{q * p:.2f}</PrezzoTotale>"
        f"<AliquotaIVA>{iva:.2f}</AliquotaIVA></DettaglioLinee>" for i, (d, q, um, p, iva) in enumerate(righe, 1))
    return (f'<?xml version="1.0" encoding="UTF-8"?><p:FatturaElettronica versione="FPR12" '
            f'xmlns:p="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"><FatturaElettronicaHeader>'
            f'<CedentePrestatore><DatiAnagrafici><IdFiscaleIVA><IdPaese>IT</IdPaese><IdCodice>{piva}</IdCodice></IdFiscaleIVA>'
            f'<Anagrafica><Denominazione>{nome}</Denominazione></Anagrafica></DatiAnagrafici></CedentePrestatore>'
            f'</FatturaElettronicaHeader><FatturaElettronicaBody><DatiGenerali><DatiGeneraliDocumento>'
            f'<TipoDocumento>TD01</TipoDocumento><Divisa>EUR</Divisa><Data>{giorno}</Data><Numero>{numero}</Numero>'
            f'</DatiGeneraliDocumento></DatiGenerali><DatiBeniServizi>{linee}</DatiBeniServizi>'
            f'</FatturaElettronicaBody></p:FatturaElettronica>')


def main():
    for d in ["dati", "db", "report", "anagrafiche"]:
        shutil.rmtree(BASE / d, ignore_errors=True)
    A, I = BASE / "anagrafiche", BASE / "dati" / "inbox"

    # ---------- anagrafiche (si compilano una volta e si aggiornano solo se cambia il menu)
    csv(A / "articoli.csv", ["codice", "descrizione", "categoria", "gruppo_me", "aliquota_iva"],
        [[c, d, cat, g, 10] for c, (d, cat, g, p, n) in ARTICOLI.items()])
    csv(A / "ricette.csv", ["codice", "ingrediente", "quantita"],
        [[c, i, float(q)] for c, r in RICETTE.items() for i, q in r.items()])
    csv(A / "ingredienti.csv", ["ingrediente", "unita", "costo_unitario_default", "classe"],
        [[i, u, float(c), CLASSE.get(i, "C")] for i, (u, c, *_) in ING.items()])
    mappa = [[FORNITORI[f][0], desc.split(" ")[0] + " " + desc.split(" ")[1], i, float(fatt),
              "merci_beverage" if f == "C" else "merci_food"] for i, (u, c, f, desc, um, fatt) in ING.items()]
    mappa += [["*", "ENERGIA", "", "", "utenze"], ["*", "GAS NATURALE", "", "", "utenze"],
              ["*", "CANONE LOCAZIONE", "", "", "affitto"], ["*", "MANUTENZIONE", "", "", "manutenzioni"]]
    csv(A / "mappa_acquisti.csv", ["piva", "contiene", "ingrediente", "fattore", "categoria_costo"], mappa)
    csv(A / "personale.csv", ["dipendente", "costo_orario"], [[n, float(c)] for n, (c, *_) in PERSONALE.items()])
    csv(A / "costi_fissi.csv", ["voce", "importo_mensile"],
        [["Affitto locale", 2200.0], ["Commercialista", 350.0], ["Assicurazioni", 180.0],
         ["Leasing forno e attrezzature", 420.0], ["Software e POS", 120.0]])

    # ---------- vendite giornaliere, coperti, presenze
    presenze, n_fatt = {}, 1
    scorta = {i: 0.0 for i in ING}  # scorta obiettivo = consumo medio di ~8 giorni
    for c, (d, cat, gr, p, media) in ARTICOLI.items():
        for ing, x in RICETTE[c].items():
            scorta[ing] += media * 1.1 * 8 * x * SPRECO.get(ing, 1.0)
    magazzino = dict(scorta)  # magazzino reale, parte pieno
    conte = {}
    inventari = {INIZIO: [[i, round(q, 2)] for i, q in magazzino.items()]}
    g = INIZIO
    while g <= FINE:
        if g.weekday() != 0:  # lunedì chiuso
            fattore = {4: 1.6, 5: 1.7, 6: 1.2}.get(g.weekday(), 0.8)
            vendite, coperti_sala, ordini_asp = [], 0, 0
            for c, (d, cat, gr, p, media) in ARTICOLI.items():
                tot = max(0, round(random.gauss(media * fattore, media * fattore * 0.2)))
                asporto = round(tot * 0.25) if cat == "Pizze" else 0
                for canale, q in (("sala", tot - asporto), ("asporto", asporto)):
                    if q:
                        nsc = round(q * 0.03)  # piccola quota non scontrinata, tracciata a parte
                        for flag, qq in (("S", q - nsc), ("N", nsc)):
                            if qq:
                                vendite.append([g.strftime("%d/%m/%Y"), canale, c, d, cat, qq, float(qq * p), flag])
                        for ing, x in RICETTE[c].items():
                            magazzino[ing] -= q * x * SPRECO.get(ing, 1.0)
                if cat == "Pizze":
                    coperti_sala += tot - asporto
                    ordini_asp += math.ceil(asporto / 2)
            csv(I / "cassa" / f"vendite_{g}.csv",
                ["data", "canale", "codice", "descrizione", "categoria", "quantita", "importo_lordo", "scontrinato"], vendite)
            csv(I / "cassa" / f"coperti_{g}.csv", ["data", "canale", "coperti"],
                [[g.strftime("%d/%m/%Y"), "sala", coperti_sala], [g.strftime("%d/%m/%Y"), "asporto", ordini_asp]])
            for n, (c, rep, ore, quando) in PERSONALE.items():
                if quando == "tutti" or g.weekday() >= 4:
                    presenze.setdefault(g.strftime("%Y-%m"), []).append(
                        [g.strftime("%d/%m/%Y"), n, float(ore + random.choice([0, 0, .5, 1])), rep])

        # ---------- ogni sabato: fatture della settimana per fornitore, a confezioni intere
        if g.weekday() == 5:
            per_forn = {}
            for ing in ING:
                u, costo, f, desc, um, fatt = ING[ing]
                conf = max(0, math.ceil((scorta[ing] - magazzino[ing]) / fatt))
                magazzino[ing] += conf * fatt
                if conf:
                    rincaro = 1.12 if (ing == "mozzarella" and g.month == 11) else random.uniform(.98, 1.04)
                    per_forn.setdefault(f, []).append((desc, float(conf), um, costo * fatt * rincaro, 10.0 if f != "C" else 22.0))
            if "A" in per_forn:
                per_forn["A"].append(("TOVAGLIOLI CARTA 2 VELI", 4.0, "CT", 18.0, 22.0))  # riga da classificare
            for f, righe in per_forn.items():
                (I / "fatture").mkdir(parents=True, exist_ok=True)
                (I / "fatture" / f"IT{FORNITORI[f][0]}_{n_fatt:05d}.xml").write_text(
                    fattura_xml(f, f"{n_fatt}/{g.year}", g, righe), encoding="utf-8")
                n_fatt += 1
        if g.day == 5:  # utenze mensili
            for desc, imp in (("FORNITURA ENERGIA ELETTRICA", 1450.0), ("FORNITURA GAS NATURALE", 620.0)):
                (I / "fatture" / f"IT09999999999_{n_fatt:05d}.xml").write_text(
                    fattura_xml("A", f"U{n_fatt}", g, [(desc, 1.0, "", imp, 22.0)]).replace(
                        FORNITORI["A"][1], "Energia Italia Spa").replace(FORNITORI["A"][0], "09999999999"), encoding="utf-8")
                n_fatt += 1
        g += timedelta(days=1)
        if g.weekday() == 0 and g.day != 1:  # conta settimanale del lunedì, solo classe A
            conte[g] = [[i, round(magazzino[i] * random.uniform(.98, 1.02), 2)] for i in CLASSE if CLASSE[i] == "A"]
        if g.day == 1:  # inventario di inizio mese = magazzino reale (con piccoli errori di conta)
            inventari[g] = [[i, round(q * random.uniform(.98, 1.02), 2)] for i, q in magazzino.items()]

    for mese, righe in presenze.items():
        csv(I / "presenze" / f"presenze_{mese}.csv", ["data", "dipendente", "ore", "reparto"], righe)

    # ---------- inventari di inizio mese
    for giorno, righe in inventari.items():
        csv(I / "inventari" / f"inventario_{giorno}.csv", ["ingrediente", "quantita"], righe)
    for giorno, righe in conte.items():
        csv(I / "conte" / f"conta_{giorno}.csv", ["ingrediente", "quantita"], righe)
    print("Dati di esempio creati. Ora lancia:  python pipeline.py")


if __name__ == "__main__":
    main()
