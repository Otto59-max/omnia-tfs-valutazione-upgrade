#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
genera_schede.py
=================
Genera il PDF della scheda di valutazione OMNIA-TFS in versione leggibile
otticamente (OMR), piu' il file template.json con la geometria delle caselle
che serve allo script di lettura.

Federazione Nazionale Maestri del Lavoro - Settore Scuola

USO
---
  # una classe, 25 copie
  python genera_schede.py --rif "MN|ITIS Fermi|3A|2026-10-14|Rossi" --copie 25 \
         --out schede_3A.pdf

  # piu' incontri in un colpo solo, da CSV (consolato;scuola;classe;data;relatore;copie)
  python genera_schede.py --da-csv incontri.csv --out schede_ottobre.pdf

  # solo il template (se hai gia' stampato le schede)
  python genera_schede.py --solo-template

DIPENDENZE
----------
  pip install reportlab qrcode pillow
"""

import argparse
import csv
import json
import os
import sys

import qrcode
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as rl_canvas

# --------------------------------------------------------------------------
# 1. CONTENUTO DEL QUESTIONARIO
#    Modificare qui (e solo qui) se la Federazione cambia le domande.
#    tipo   : "singola" = una sola risposta ammessa; "multipla" = piu' risposte
#    scala  : punteggio 1-5 associato a ogni opzione, per l'indice sintetico.
#             None se la domanda non e' una scala di gradimento.
# --------------------------------------------------------------------------
DOMANDE = [
    {
        "id": "D01",
        "testo": "Gli argomenti trattati ti sono sembrati interessanti:",
        "tipo": "singola",
        "scala": [5, 4, 3, 2, 1],
        "opzioni": ["Moltissimo", "Molto", "Abbastanza", "Poco", "Per niente"],
    },
    {
        "id": "D02",
        "testo": "Il tempo dedicato agli argomenti \u00e8 stato:",
        "tipo": "singola",
        "scala": None,
        "opzioni": ["Troppo limitato", "Insufficiente", "Giusto",
                    "Abbondante", "Troppo lungo"],
    },
    {
        "id": "D03",
        "testo": "Il relatore/i \u00e8 stato chiaro nelle spiegazioni e nei contenuti:",
        "tipo": "singola",
        "scala": [5, 4, 3, 2, 1],
        "opzioni": ["Eccellente", "Molto", "Abbastanza", "Poco", "Per niente"],
    },
    {
        "id": "D04",
        "testo": "Tra le attivit\u00e0 proposte quali ti sono piaciute di pi\u00f9:",
        "tipo": "multipla",
        "scala": None,
        "opzioni": ["Racconti di episodi di lavoro", "Lavori di gruppo",
                    "Filmati - Video", "Interviste", "Esempi"],
    },
    {
        "id": "D05",
        "testo": "L'incontro ti ha fatto riflettere sul tuo futuro anche scolastico:",
        "tipo": "singola",
        "scala": [5, 4, 3, 2, 1],
        "opzioni": ["Decisamente s\u00ec", "Molto", "Abbastanza", "Poco", "Per niente"],
    },
    {
        "id": "D06",
        "testo": "Pensi che quello che hai ascoltato ti potr\u00e0 essere utile:",
        "tipo": "singola",
        "scala": [5, 4, 3, 2, 1],
        "opzioni": ["Moltissimo", "Molto", "Abbastanza", "Poco", "Per niente"],
    },
    {
        "id": "D07",
        "testo": "Come ti sei sentito partecipando all'incontro:",
        "tipo": "singola",
        "scala": None,
        "opzioni": ["Coinvolto", "Interessato", "Curioso", "Confuso", "Annoiato"],
    },
    {
        "id": "D08",
        "testo": "Ritieni utile alla tua formazione gli incontri con i Maestri del Lavoro:",
        "tipo": "singola",
        "scala": [5, 4, 3, 2, 1],
        "opzioni": ["Moltissimo", "Molto", "Abbastanza", "Poco", "Per niente"],
    },
    {
        "id": "D09",
        "testo": "Complessivamente, cosa ti \u00e8 piaciuto di pi\u00f9 dell'incontro:",
        "tipo": "multipla",
        "scala": None,
        "opzioni": ["Lo stile espositivo",
                    "L'utilizzo di strumenti digitali (video, etc.)",
                    "I momenti di confronto, dibattito",
                    "Le attivit\u00e0 di gruppo",
                    "I temi legati al mio futuro e al lavoro",
                    "Non so dare un giudizio"],
    },
    {
        "id": "D10",
        "testo": "Quali miglioramenti suggeriresti nei prossimi incontri:",
        "tipo": "multipla",
        "scala": None,
        "opzioni": ["Maggiore interazione con gli studenti",
                    "Pi\u00f9 attivit\u00e0 pratiche",
                    "Maggiore allineamento con attualit\u00e0 - futuro",
                    "Migliorare il materiale espositivo",
                    "Pi\u00f9 esperienze lavorative",
                    "Non cambierei nulla, va bene cos\u00ec"],
    },
]

# --------------------------------------------------------------------------
# 2. PARAMETRI DI LAYOUT (in millimetri)
#    I valori qui sotto riproducono il layout approvato (Layout_Scheda.pdf):
#    intestazione a 25,5 mm, griglia a 60,8 mm, passo opzioni 5,0 mm,
#    nessun riquadro commenti.
# --------------------------------------------------------------------------
PAG_W, PAG_H = A4                     # punti
PAG_W_MM, PAG_H_MM = 210.0, 297.0

MARGINE = 10.0                        # margine esterno
MARKER_LATO = 6.0                     # lato dei quadrati di riferimento
MARKER_OFF = MARGINE + MARKER_LATO / 2.0   # centro marker dal bordo

BOLLA_R = 2.3                         # raggio casella da annerire
BOLLA_SPESSORE = 0.9                  # spessore bordo casella (punti)

# Zona di rispetto: nessun elemento stampato vicino ai marker, altrimenti in
# scansione i contorni si fondono e il raddrizzamento salta.
CONT_X0 = 16.0                        # colonna di testo sinistra
CONT_X1 = 194.0                       # bordo destro del contenuto
CONT_Y1 = PAG_H_MM - MARGINE - MARKER_LATO - 3.5   # 277,5 mm

Y_TESTATA = 25.5                      # alto del blocco intestazione + QR
H_TESTATA = 19.0                      # altezza del QR, definisce il blocco
Y_ISTRUZIONI = 50.0                   # riga "come compilare"
Y_GRIGLIA = 60.8                      # alto del primo riquadro domanda

COL_GAP = 5.0                         # spazio fra le due colonne
RIGA_OPZ = 5.0                        # passo verticale fra le opzioni
RIGA_OPZ_MIN = 4.2                    # limite di compressione automatica
PAD_BOX = 2.6                         # padding interno ai riquadri domanda
GAP_BOX = 2.4                         # spazio fra un riquadro e il successivo
F_TESTO = 7.6                         # corpo del testo di domande e opzioni
H_COMMENTI = 0.0                      # 0 = nessun riquadro commenti

F_TITOLO = 12.5

BLU = HexColor("#1F6FB2")
GRIGIO = HexColor("#9A9A9A")


# --------------------------------------------------------------------------
# 3. UTILITA' DI DISEGNO
#    Origine in alto a sinistra; y_pdf() converte nel sistema di reportlab.
# --------------------------------------------------------------------------
def y_pdf(y_mm):
    """Converte una y misurata dall'alto in coordinate reportlab."""
    return (PAG_H_MM - y_mm) * mm


def spezza(testo, font, size, larghezza_mm):
    """Manda a capo il testo entro la larghezza data. Ritorna lista di righe."""
    limite = larghezza_mm * mm
    parole = testo.split()
    righe, corrente = [], ""
    for p in parole:
        prova = (corrente + " " + p).strip()
        if stringWidth(prova, font, size) <= limite:
            corrente = prova
        else:
            if corrente:
                righe.append(corrente)
            corrente = p
    if corrente:
        righe.append(corrente)
    return righe or [""]


NOTA_SINGOLA = "(una sola risposta)"
NOTA_MULTIPLA = "(anche pi\u00f9 di una)"


def _nota(dom):
    return NOTA_SINGOLA if dom["tipo"] == "singola" else NOTA_MULTIPLA


def righe_domanda(dom, numero, larghezza_mm, f_testo):
    """Righe del testo della domanda e posizione della nota.

    La nota ("una sola risposta") va in coda all'ultima riga se ci sta,
    altrimenti su una riga propria.
    """
    utile = larghezza_mm - 2 * PAD_BOX
    righe = spezza("%d. %s" % (numero, dom["testo"]), "Helvetica-Bold",
                   f_testo, utile)
    w_ultima = stringWidth(righe[-1], "Helvetica-Bold", f_testo)
    w_nota = stringWidth(_nota(dom), "Helvetica-Oblique", 6.6)
    inline = (w_ultima + 2.2 * mm + w_nota) <= utile * mm
    return righe, inline


def altezza_domanda(dom, numero, larghezza_mm, riga_opz, f_testo):
    """Altezza del riquadro di una domanda, in mm."""
    righe, inline = righe_domanda(dom, numero, larghezza_mm, f_testo)
    h = PAD_BOX + len(righe) * (f_testo * 0.352778 + 1.1)
    if not inline:
        h += 6.6 * 0.352778 + 1.0
    h += 1.4 + len(dom["opzioni"]) * riga_opz + PAD_BOX
    return h


def calcola_layout():
    """Parametri di griglia. Comprime solo se il contenuto non entra.

    Con le domande attuali restituisce i valori nominali, che riproducono
    il layout approvato. Se un giorno si aggiungono domande o si allungano
    le opzioni, il passo verticale si riduce da solo invece di sbordare
    sulla seconda pagina.
    """
    col_w = (CONT_X1 - CONT_X0 - COL_GAP) / 2.0
    disponibile = CONT_Y1 - Y_GRIGLIA - (H_COMMENTI + 2.5 if H_COMMENTI else 0.0)

    passo = RIGA_OPZ
    while passo >= RIGA_OPZ_MIN - 0.001:
        for f_testo in (F_TESTO, 7.2, 6.8):
            altezze = [altezza_domanda(d, i + 1, col_w, passo, f_testo)
                       for i, d in enumerate(DOMANDE)]
            totale = sum(max(altezze[r:r + 2]) + GAP_BOX
                         for r in range(0, len(DOMANDE), 2))
            if totale - GAP_BOX <= disponibile:
                return {"riga_opz": passo, "f_testo": f_testo, "col_w": col_w}
        passo -= 0.1
    return {"riga_opz": RIGA_OPZ_MIN, "f_testo": 6.8, "col_w": col_w}


def disegna_marker(c):
    """I quattro quadrati neri pieni che permettono di raddrizzare la scansione.

    Sono pieni e senza fori: cosi' lo script di lettura li distingue dai
    quadrati di ricerca del QR code, che hanno un anello bianco interno.
    """
    centri = [
        (MARKER_OFF, MARKER_OFF),
        (PAG_W_MM - MARKER_OFF, MARKER_OFF),
        (PAG_W_MM - MARKER_OFF, PAG_H_MM - MARKER_OFF),
        (MARKER_OFF, PAG_H_MM - MARKER_OFF),
    ]
    c.setFillColor(black)
    for cx, cy in centri:
        c.rect((cx - MARKER_LATO / 2) * mm, y_pdf(cy + MARKER_LATO / 2),
               MARKER_LATO * mm, MARKER_LATO * mm, stroke=0, fill=1)
    return centri


def qr_immagine(payload):
    """Genera il QR come oggetto PIL."""
    qr = qrcode.QRCode(version=None,
                       error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=10, border=2)
    qr.add_data(payload)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("RGB")


# --------------------------------------------------------------------------
# 4. DISEGNO DI UNA SCHEDA
# --------------------------------------------------------------------------
def disegna_scheda(c, payload, etichetta_leggibile, lay):
    from reportlab.lib.utils import ImageReader

    geometria = {"domande": [], "qr_zona": None, "commenti_zona": None}

    disegna_marker(c)
    rif_x0, rif_y0 = MARKER_OFF, MARKER_OFF
    rif_w = PAG_W_MM - 2 * MARKER_OFF
    rif_h = PAG_H_MM - 2 * MARKER_OFF

    def norm(x_mm, y_mm):
        """Da millimetri-pagina a coordinate normalizzate sul rettangolo marker."""
        return round((x_mm - rif_x0) / rif_w, 6), round((y_mm - rif_y0) / rif_h, 6)

    x0, x1 = CONT_X0, CONT_X1
    f_testo = lay["f_testo"]
    riga_opz = lay["riga_opz"]

    # ---------------- intestazione ----------------
    y = Y_TESTATA
    qr_lato = H_TESTATA
    qr_x = x1 - qr_lato
    c.drawImage(ImageReader(qr_immagine(payload)),
                qr_x * mm, y_pdf(y + qr_lato), qr_lato * mm, qr_lato * mm)
    u0, v0 = norm(qr_x - 2.5, y - 2.5)
    u1, v1 = norm(qr_x + qr_lato + 2.5, y + qr_lato + 2.5)
    geometria["qr_zona"] = {"u0": u0, "v0": v0, "u1": u1, "v1": v1}

    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(x0 * mm, y_pdf(y + 3.6), "FEDERAZIONE NAZIONALE")
    c.drawString(x0 * mm, y_pdf(y + 7.6), "MAESTRI DEL LAVORO")

    c.setFillColor(BLU)
    c.setFont("Helvetica-Bold", F_TITOLO)
    c.drawString(x0 * mm, y_pdf(y + 14.6),
                 "QUESTIONARIO VALUTAZIONE INCONTRO \u2013 OMNIA TFS")

    c.setFillColor(GRIGIO)
    c.setFont("Helvetica", 7.2)
    c.drawString(x0 * mm, y_pdf(y + 18.6), "RIFERIMENTO: " + etichetta_leggibile)

    # ---------------- istruzioni ----------------
    c.setFillColor(black)
    c.circle((x0 + 2.4) * mm, y_pdf(Y_ISTRUZIONI + 2.2), BOLLA_R * mm,
             stroke=1, fill=1)
    c.setFont("Helvetica-Bold", 7.6)
    c.drawString((x0 + 6.2) * mm, y_pdf(Y_ISTRUZIONI + 2.9),
                 "COME COMPILARE:  annerisci completamente la casella con penna "
                 "nera o blu.  Non usare la matita.")

    # ---------------- griglia domande, 2 colonne ----------------
    col_w = lay["col_w"]
    col_x = [x0, x0 + col_w + COL_GAP]

    # Le due colonne sono allineate per riga: l'altezza di ogni riga e' quella
    # del suo riquadro piu' alto, cosi' la scheda resta una griglia leggibile
    # invece di due colonne che scorrono indipendenti.
    altezze = [altezza_domanda(d, i + 1, col_w, riga_opz, f_testo)
               for i, d in enumerate(DOMANDE)]
    y_riga, quota = [], Y_GRIGLIA
    for r in range(0, len(DOMANDE), 2):
        y_riga.append(quota)
        quota += max(altezze[r:r + 2]) + GAP_BOX

    for idx, dom in enumerate(DOMANDE):
        col = idx % 2
        bx, by = col_x[col], y_riga[idx // 2]
        h = max(altezze[idx - col:idx - col + 2])

        c.setStrokeColor(HexColor("#B4B4B4"))
        c.setLineWidth(0.6)
        c.rect(bx * mm, y_pdf(by + h), col_w * mm, h * mm, stroke=1, fill=0)

        righe, inline = righe_domanda(dom, idx + 1, col_w, f_testo)
        h_riga = f_testo * 0.352778 + 1.1
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", f_testo)
        ty = by + PAD_BOX + f_testo * 0.352778
        for k, r in enumerate(righe):
            c.drawString((bx + PAD_BOX) * mm, y_pdf(ty), r)
            if k == len(righe) - 1 and inline:
                c.setFont("Helvetica-Oblique", 6.6)
                c.setFillColor(GRIGIO)
                c.drawString((bx + PAD_BOX) * mm +
                             stringWidth(r, "Helvetica-Bold", f_testo) + 2.2 * mm,
                             y_pdf(ty), _nota(dom))
                c.setFillColor(black)
                c.setFont("Helvetica-Bold", f_testo)
            ty += h_riga
        if not inline:
            c.setFont("Helvetica-Oblique", 6.6)
            c.setFillColor(GRIGIO)
            c.drawString((bx + PAD_BOX) * mm, y_pdf(ty), _nota(dom))
            ty += 6.6 * 0.352778 + 1.0
        ty += 1.4

        geo_opz = []
        for j, opz in enumerate(dom["opzioni"]):
            cy = ty + j * riga_opz + riga_opz / 2.0 - 1.0
            cx = bx + PAD_BOX + BOLLA_R
            c.setStrokeColor(black)
            c.setFillColor(white)
            c.setLineWidth(BOLLA_SPESSORE)
            c.circle(cx * mm, y_pdf(cy), BOLLA_R * mm, stroke=1, fill=0)
            c.setFillColor(black)
            c.setFont("Helvetica", f_testo)
            c.drawString((cx + BOLLA_R + 2.0) * mm, y_pdf(cy + 1.2), opz)
            u, v = norm(cx, cy)
            geo_opz.append({"indice": j, "etichetta": opz, "u": u, "v": v})

        geometria["domande"].append({
            "id": dom["id"], "numero": idx + 1, "testo": dom["testo"],
            "tipo": dom["tipo"], "scala": dom["scala"], "opzioni": geo_opz})

    # ---------------- area commenti (solo se H_COMMENTI > 0) ----------------
    if H_COMMENTI > 0:
        y = min(quota + 0.5, CONT_Y1 - H_COMMENTI)
        c.setStrokeColor(HexColor("#B4B4B4"))
        c.setLineWidth(0.6)
        c.rect(x0 * mm, y_pdf(y + H_COMMENTI), (x1 - x0) * mm, H_COMMENTI * mm,
               stroke=1, fill=0)
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", f_testo)
        c.drawString((x0 + PAD_BOX) * mm, y_pdf(y + PAD_BOX + 2.4),
                     "Commenti facoltativi:")
        cu0, cv0 = norm(x0, y)
        cu1, cv1 = norm(x1, y + H_COMMENTI)
        geometria["commenti_zona"] = {"u0": cu0, "v0": cv0, "u1": cu1, "v1": cv1}

    geometria["rettangolo_marker_mm"] = {
        "x0": rif_x0, "y0": rif_y0, "larghezza": rif_w, "altezza": rif_h}
    geometria["raggio_bolla_mm"] = BOLLA_R
    return geometria


# --------------------------------------------------------------------------
# 5. CALIBRAZIONE DA UN PDF ESISTENTE
#    Ricava template.json leggendo direttamente marker, caselle e QR dal
#    contenuto vettoriale di un PDF di layout, chiunque lo abbia prodotto.
#    Serve quando la scheda viene ritoccata fuori da questo script: il
#    lettore resta allineato senza dover indovinare le coordinate.
# --------------------------------------------------------------------------
def calibra_da_pdf(percorso):
    import pymupdf

    MM_PT = 72.0 / 25.4
    pagina = pymupdf.open(percorso)[0]

    marker, riquadri, bolle = [], [], []
    for tratto in pagina.get_drawings():
        pieno = tratto.get("fill") is not None
        for elem in tratto["items"]:
            if elem[0] != "re":
                continue
            r = elem[1]
            x, y = r.x0 / MM_PT, r.y0 / MM_PT
            w, h = r.width / MM_PT, r.height / MM_PT
            if pieno and abs(w - MARKER_LATO) < 1.0 and abs(h - MARKER_LATO) < 1.0:
                marker.append((x + w / 2.0, y + h / 2.0))
            elif not pieno and w > 20.0 and h > 5.0:
                riquadri.append((x, y, x + w, y + h))
        if tratto["items"] and all(e[0] == "c" for e in tratto["items"]):
            r = tratto["rect"]
            bolle.append(((r.x0 + r.x1) / 2.0 / MM_PT,
                          (r.y0 + r.y1) / 2.0 / MM_PT,
                          r.width / 2.0 / MM_PT))

    if len(marker) != 4:
        raise ValueError("trovati %d marker invece di 4: il PDF non e' una "
                         "scheda OMR valida" % len(marker))

    xs = sorted(c[0] for c in marker)
    ys = sorted(c[1] for c in marker)
    rif_x0, rif_y0 = (xs[0] + xs[1]) / 2.0, (ys[0] + ys[1]) / 2.0
    rif_w = (xs[2] + xs[3]) / 2.0 - rif_x0
    rif_h = (ys[2] + ys[3]) / 2.0 - rif_y0

    def norm(x_mm, y_mm):
        return round((x_mm - rif_x0) / rif_w, 6), round((y_mm - rif_y0) / rif_h, 6)

    def dentro(b, q):
        return q[0] - 0.5 <= b[0] <= q[2] + 0.5 and q[1] - 0.5 <= b[1] <= q[3] + 0.5

    # righe della griglia: raggruppa i riquadri per quota, poi ordina da sinistra
    riquadri.sort(key=lambda q: q[1])
    gruppi, corrente = [], []
    for q in riquadri:
        if corrente and q[1] - corrente[0][1] > 10.0:
            gruppi.append(sorted(corrente, key=lambda t: t[0]))
            corrente = []
        corrente.append(q)
    if corrente:
        gruppi.append(sorted(corrente, key=lambda t: t[0]))
    ordinati = [q for g in gruppi for q in g]

    con_bolle = [(q, sorted([b for b in bolle if dentro(b, q)], key=lambda t: t[1]))
                 for q in ordinati]
    domande_box = [(q, bs) for q, bs in con_bolle if bs]
    senza_bolle = [q for q, bs in con_bolle if not bs]

    if len(domande_box) != len(DOMANDE):
        raise ValueError("il PDF contiene %d riquadri-domanda, le domande "
                         "definite sono %d" % (len(domande_box), len(DOMANDE)))

    geometria = {"domande": [], "qr_zona": None, "commenti_zona": None}
    for idx, (dom, (_q, bs)) in enumerate(zip(DOMANDE, domande_box)):
        if len(bs) != len(dom["opzioni"]):
            raise ValueError("domanda %d: nel PDF ci sono %d caselle, le opzioni "
                             "definite sono %d" % (idx + 1, len(bs),
                                                   len(dom["opzioni"])))
        opz = []
        for j, (bx, by, _r) in enumerate(bs):
            u, v = norm(bx, by)
            opz.append({"indice": j, "etichetta": dom["opzioni"][j],
                        "u": u, "v": v})
        geometria["domande"].append({
            "id": dom["id"], "numero": idx + 1, "testo": dom["testo"],
            "tipo": dom["tipo"], "scala": dom["scala"], "opzioni": opz})

    if senza_bolle:
        q = max(senza_bolle, key=lambda t: (t[3] - t[1]) * (t[2] - t[0]))
        u0, v0 = norm(q[0], q[1])
        u1, v1 = norm(q[2], q[3])
        geometria["commenti_zona"] = {"u0": u0, "v0": v0, "u1": u1, "v1": v1}

    rett_qr = [r for im in pagina.get_images(full=True)
               for r in pagina.get_image_rects(im[0])]
    if rett_qr:
        r = max(rett_qr, key=lambda t: t.width * t.height)
        u0, v0 = norm(r.x0 / MM_PT - 2.5, r.y0 / MM_PT - 2.5)
        u1, v1 = norm(r.x1 / MM_PT + 2.5, r.y1 / MM_PT + 2.5)
        geometria["qr_zona"] = {"u0": u0, "v0": v0, "u1": u1, "v1": v1}

    raggi = sorted(b[2] for _q, bs in domande_box for b in bs)
    geometria["raggio_bolla_mm"] = round(raggi[len(raggi) // 2], 3)
    geometria["rettangolo_marker_mm"] = {
        "x0": rif_x0, "y0": rif_y0, "larghezza": rif_w, "altezza": rif_h}
    geometria["calibrato_da"] = os.path.basename(percorso)
    return geometria


# --------------------------------------------------------------------------
# 5. MAIN
# --------------------------------------------------------------------------
def payload_da_campi(consolato, scuola, classe, data, relatore):
    """Contenuto del QR: campi separati da '|', prefisso di riconoscimento."""
    campi = [str(x).replace("|", "/").strip()
             for x in (consolato, scuola, classe, data, relatore)]
    return "OMNIA|" + "|".join(campi)


def etichetta_da_payload(payload):
    p = payload.split("|")
    return "  \u00b7  ".join(x for x in p[1:] if x)


def main():
    ap = argparse.ArgumentParser(
        description="Genera le schede OMNIA-TFS in versione a lettura ottica.")
    ap.add_argument("--rif", help="Riferimento singolo: "
                    "\"consolato|scuola|classe|data|relatore\"")
    ap.add_argument("--copie", type=int, default=1,
                    help="Numero di copie per il riferimento singolo")
    ap.add_argument("--da-csv", help="CSV con colonne "
                    "consolato;scuola;classe;data;relatore;copie")
    ap.add_argument("--out", default="schede.pdf", help="PDF di uscita")
    ap.add_argument("--template", default="template.json",
                    help="File di geometria per lo script di lettura")
    ap.add_argument("--solo-template", action="store_true",
                    help="Genera solo il template.json, senza PDF stampabile")
    ap.add_argument("--calibra-da", metavar="PDF",
                    help="Ricava il template da un PDF di layout gia' pronto "
                         "(usare quando la scheda e' stata ritoccata a mano)")
    args = ap.parse_args()

    if args.calibra_da:
        geometria = calibra_da_pdf(args.calibra_da)
        geometria["versione"] = 1
        geometria["pagina_mm"] = {"larghezza": PAG_W_MM, "altezza": PAG_H_MM}
        with open(args.template, "w", encoding="utf-8") as f:
            json.dump(geometria, f, ensure_ascii=False, indent=1)
        print("Template calibrato su %s e salvato in %s"
              % (args.calibra_da, args.template))
        print("  domande       : %d" % len(geometria["domande"]))
        print("  caselle       : %d"
              % sum(len(d["opzioni"]) for d in geometria["domande"]))
        print("  raggio casella: %.2f mm" % geometria["raggio_bolla_mm"])
        print("  QR            : %s"
              % ("rilevato" if geometria["qr_zona"] else "assente"))
        print("  commenti      : %s"
              % ("rilevati" if geometria["commenti_zona"] else "assenti"))
        return 0

    lotti = []
    if args.da_csv:
        with open(args.da_csv, newline="", encoding="utf-8-sig") as f:
            for riga in csv.DictReader(f, delimiter=";"):
                lotti.append((
                    payload_da_campi(riga.get("consolato", ""), riga.get("scuola", ""),
                                     riga.get("classe", ""), riga.get("data", ""),
                                     riga.get("relatore", "")),
                    int(riga.get("copie", 1) or 1)))
    elif args.rif:
        campi = (args.rif.split("|") + [""] * 5)[:5]
        lotti.append((payload_da_campi(*campi), max(1, args.copie)))
    else:
        lotti.append((payload_da_campi("PROVA", "Scuola di prova", "3A",
                                       "2026-10-14", "Rossi"), 1))

    out = args.out if not args.solo_template else os.devnull
    c = rl_canvas.Canvas(out, pagesize=A4)
    c.setTitle("Questionario valutazione incontro - OMNIA TFS")
    lay = calcola_layout()

    geometria = None
    totale = 0
    for payload, copie in lotti:
        n = 1 if args.solo_template else copie
        for _ in range(n):
            g = disegna_scheda(c, payload, etichetta_da_payload(payload), lay)
            geometria = geometria or g
            c.showPage()
            totale += 1
        if args.solo_template:
            break
    c.save()

    geometria["versione"] = 1
    geometria["pagina_mm"] = {"larghezza": PAG_W_MM, "altezza": PAG_H_MM}
    with open(args.template, "w", encoding="utf-8") as f:
        json.dump(geometria, f, ensure_ascii=False, indent=1)

    if args.solo_template:
        print("Template salvato in %s" % args.template)
    else:
        print("Generate %d schede in %s" % (totale, args.out))
        print("Template geometrico salvato in %s (serve a leggi_schede.py)"
              % args.template)


if __name__ == "__main__":
    sys.exit(main())
