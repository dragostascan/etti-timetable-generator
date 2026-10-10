import os
import sys
import difflib
import pandas as pd
import re
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ==========================================
# 1. CONFIGURARE GLOBALA
# ==========================================
# Aici definim dicționarele de traducere. Secretariatul scrie în "Anunțuri" numele lung (ex: "Fizică"),
# dar în grila de "Orar" scrie prescurtat ("fiz2 (l)"). Aceste dicționare fac legătura între ele.

TRADUCATOR_RO = {
    
    # ================= ANUL 1 =================
    "fiz2 (l)": "Fizică",
    "fiz2(l)": "Fizică",
    "pclp2 (l)": "Programarea Calculatoarelor și Limbaje de Programare 2 (PCLP)",
    "pclp2(l)": "Programarea Calculatoarelor și Limbaje de Programare 2 (PCLP)",
    "metc (l)": "Măsurări în Electronică și Telecomunicații (METc)",
    "metc(l)": "Măsurări în Electronică și Telecomunicații (METc)",
    "ia (p)": "Informatica aplicata - proiect",
    "ia(p)": "Informatica aplicata - proiect",
    "so1 (l)": "Sisteme de Operare 1 (SO1)",
    "so1(l)": "Sisteme de Operare 1 (SO1)",
    "me (l)": "Materiale pentru Electronică (ME)",
    "me(l)": "Materiale pentru Electronică (ME)",
    "sport (s)": "Educaţie fizică şi sport 2",
    "sport(s)": "Educaţie fizică şi sport 2",
    "sport": "Educaţie fizică şi sport 2",

    # ================= ANUL 2 =================
    "amp (l)": "Arhitectura microprocesoarelor 2. Microcontrolere (laborator)",
    "amp(l)": "Arhitectura microprocesoarelor 2. Microcontrolere (laborator)",
    "bd (l)": "Baze de date (laborator)",
    "bd(l)": "Baze de date (laborator)",
    "cef (l)": "Circuite Electronice Fundamentale + DEEA2 - Laborator",
    "cef(l)": "Circuite Electronice Fundamentale + DEEA2 - Laborator",
    "cid (l)": "Circuite Integrate Digitale + ED - Laborator",
    "cid(l)": "Circuite Integrate Digitale + ED - Laborator",
    "ss (l)": "Semnale și sisteme 2 (laborator)",
    "ss(l)": "Semnale și sisteme 2 (laborator)",
    "ss2 (l)": "Semnale și sisteme 2 (laborator)",
    "ss2(l)": "Semnale și sisteme 2 (laborator)",
    "sp (s)": "Sport (SP) (seminar)",
    "sp(s)": "Sport (SP) (seminar)",
    "ma2 (l)": "Matematici Aplicate 2 (laborator)",
    "ma2(l)": "Matematici Aplicate 2 (laborator)",

    # ================= ANUL 3 =================
    "depi (l)": "Decizie si estimare in prelucrarea informatiilor",
    "depi(l)": "Decizie si estimare in prelucrarea informatiilor",
    "depi (s)": "Decizie si estimare in prelucrarea informatiilor",
    "depi(s)": "Decizie si estimare in prelucrarea informatiilor",
    "pds (l)": "Prelucrarea digitală a semnalelor",
    "pds(l)": "Prelucrarea digitală a semnalelor",
    "pds (s)": "Prelucrarea digitală a semnalelor",
    "pds(s)": "Prelucrarea digitală a semnalelor",
    "tv (l)": "Televiziune",
    "tv(l)": "Televiziune",
    "tpi (l)": "Tehnologia si prelucrarea informatiei",
    "tpi(l)": "Tehnologia si prelucrarea informatiei",
    "bsad (l)": "Baze de date pentru stocarea si analiza datelor",
    "bsad(l)": "Baze de date pentru stocarea si analiza datelor",
    "mic (l)": "Microprocesoare (Microprocessors)",
    "mic(l)": "Microprocesoare (Microprocessors)",
    "pr apm": "Aplicatii practice ale microcontrolerelor - Proiect",
    "uU (l)": "Microunde",
    "uU(l)": "Microunde",

    # ================= ANUL 4 =================
    "cm (l)": "COMUNICAŢII MOBILE",
    "cm(l)": "COMUNICAŢII MOBILE",
    "tcsm (l)": "TEHNICI DE COMPRESIE A SEMNALELOR MULTIMEDIA",
    "tcsm(l)": "TEHNICI DE COMPRESIE A SEMNALELOR MULTIMEDIA",
    "caf (l)": "CALITATE ŞI FIABILITATE",
    "caf(l)": "CALITATE ŞI FIABILITATE",
    "rcm (l)": "REŢELE DE COMUNICAŢII MOBILE",
    "rcm(l)": "REŢELE DE COMUNICAŢII MOBILE",
    "tstm (l)": "TEHNICI ŞI SISTEME DE TRANSMISIUNI MULTIPLEX",
    "tstm(l)": "TEHNICI ŞI SISTEME DE TRANSMISIUNI MULTIPLEX",
    "so (l)": "Sisteme de operare",
    "so(l)": "Sisteme de operare",
    "ici (l)": "Inteligenta computationala integrata",
    "ici(l)": "Inteligenta computationala integrata",
    "robo (l)": "Robotica",
    "robo(l)": "Robotica",
    "src (l)": "Sisteme reconfigurabile de calcul",
    "src(l)": "Sisteme reconfigurabile de calcul",
    "aaccep (l)": "Analiza asist. de calculator a circ. electr. de putere",
    "aaccep(l)": "Analiza asist. de calculator a circ. electr. de putere",
    "taep (l)": "Testarea automata a echipamentelor si proceselor",
    "taep(l)": "Testarea automata a echipamentelor si proceselor",
    "scm (l)": "SISTEME DE COMUNICAȚII MOBILE",
    "scm(l)": "SISTEME DE COMUNICAȚII MOBILE",
    "AI (l)": "Analiza imaginilor",
    "AI(l)": "Analiza imaginilor",
    "APD(l)": "Algoritmi paraleli si distribuiti 2",
    "APD (l)": "Algoritmi paraleli si distribuiti 2",
    "DDM(l)": "Dispozitive dielectrice si magnetice",
    "DDM (l)": "Dispozitive dielectrice si magnetice"
}

TRADUCATOR_EN = {
    # ================= ANUL 1 =================
    # Prescurtarile din orar sunt in engleza, dar in Anunturi scrie in romana
    "phyiscs2 (l)": "Fizică",
    "physics2(l)": "Fizică",
    "cppl2 (l)": "Programarea Calculatoarelor și Limbaje de Programare 2 (PCLP)",
    "cppl2(l)": "Programarea Calculatoarelor și Limbaje de Programare 2 (PCLP)",
    "metc (l)": "Măsurări în Electronică și Telecomunicații (METc)",
    "metc(l)": "Măsurări în Electronică și Telecomunicații (METc)",
    "ai (p)": "Informatica aplicata - proiect",
    "ai(p)": "Informatica aplicata - proiect",
    "em (l)": "Materiale pentru Electronică (ME)",
    "em(l)": "Materiale pentru Electronică (ME)",
    "sport": "Educaţie fizică şi sport 2",

    # ================= ANUL 2 =================
    # Aici anunturile incep sa aiba denumirile in engleza, iar orarul are prescurtari EN (ex: FEC)
    "ma2 (l)": "Microprocessor Architecture 2 (lab)",
    "ma2(l)": "Microprocessor Architecture 2 (lab)",
    "db (l)": "Data bases (laborator)",
    "db(l)": "Data bases (laborator)",
    "fec (l)": "Fundamental Electronic Circuits - Laboratory",
    "fec(l)": "Fundamental Electronic Circuits - Laboratory",
    "dic (l)": "Digital Integrated Circuits - Laboratory",
    "dic(l)": "Digital Integrated Circuits - Laboratory",
    "ss (l)": "Signals and systems 2 (lab)",
    "ss(l)": "Signals and systems 2 (lab)",
    "sport": "Sport (SP) (seminar)",

    # ================= ANUL 3 =================
    "deip (l)": "Decizie si estimare in prelucrarea informatiilor",
    "deip(l)": "Decizie si estimare in prelucrarea informatiilor",
    "dsp (l) SC2": "Prelucrarea digitală a semnalelor",
    "dsp(l) SC2": "Prelucrarea digitală a semnalelor",
    "mc(l)": "Circuite de microunde",
    "mc (l)": "Circuite de microunde",
    "tv (l)": "Televiziune",
    "tv(l)": "Televiziune",
    "nai (l) Sc.2": "Network Architectures and Internet",
    "bsad(l)": "Bazele sistemelor de achiziție de date",
    "bsad (l)": "Bazele sistemelor de achiziție de date",
    "tdav(l)": "Tehnici digitale audio video",
    "tdav (l)": "Tehnici digitale audio video",
    "ia (l) SC2": "ARTIFICIAL INTELLIGENCE",
    "ia(l) SC2": "ARTIFICIAL INTELLIGENCE",
    "uU (l)": "Microunde",
    "uU(l)": "Microunde",
    
    # ================= ANUL 4 =================
    "robo (l)": "Robotica",
    "robo(l)": "Robotica",
    "mcs (l)": "Mobile communications systems",
    "mcs(l)": "Mobile cummunications systems",
    "CAAPEC (l)": "Analiza asist. de calculator a circ. electr. de putere",
    "CAAPEC(l)": "Analiza asist. de calculator a circ. electr. de putere",
    "CAF(l)": "CALITATE ŞI FIABILITATE",
    "CAF (l)": "CALITATE ŞI FIABILITATE",
    "MC (l)": "MOBILE COMMUNICATIONS",
    "MC(l)": "MOBILE COMMUNICATIONS",
    "MC-TA (l)": "MULTIMEDIA CODING - TECHNIQUES AND APPLICATIONS",
    "MC-TA(l)": "MULTIMEDIA CODING - TECHNIQUES AND APPLICATIONS",
    "ATEP (l)": "Automatic testing of equipments and processes",
    "ATEP(l)": "Automatic testing of equipments and processes",
    "rcs (l)": "Sisteme reconfigurabile de calcul",
    "rcs(l)": "Sisteme reconfigurabile de calcul",
    "QR (l) SC1": "Quality and Reliability",
    "QR(l) SC1": "Quality and Reliability"
}

ZILE_SAPTAMANA = ['LUNI', 'MARTI', 'MIERCURI', 'JOI', 'VINERI']
# Scurtatura pt a genera lista de ore: ['09-10', '10-11', ... , '20-21']
ORE_STANDARD = [f"{i:02d}-{i+1:02d}" for i in range(9, 21)]

# ==========================================
# 2. FUNCTII UTILE (Curatare si Procesare logica)
# ==========================================
def extrage_grupa_si_anul():
    # Verifica daca grupa respecta formatul facultatii (ex: 432A -> Anul 3, Seria A, Grupa 2)
    grupa = input("Te rog introdu grupa ta (ex: 422A, 414C, 411G): ").strip().upper()
    if len(grupa) == 4 and grupa[:3].isdigit() and grupa[3].isalpha() and 1 <= int(grupa[1]) <= 4:
        return grupa, grupa[1] 
    else:
        print("Eroare: Formatul grupei este invalid! Te rog sa scrii o grupa reala (ex: 411C, 432A).")
        return extrage_grupa_si_anul()

def este_materie(text):
    # Opreste anumite formatari din Excel (cratime, spatii goale, "nan")
    t = str(text).strip()
    if re.match(r'^-+$', t): # Daca are doar cratime (---)
        return False
    return t not in ["", "Liber", "nan"]

def normalize_text(text):
    # Doar taie spatiile invizibile de la margini
    return str(text).strip()

def normalize_key(text):
    # Face toate literele mici, sterge diacriticele si spatiile duble, face legatura cu Traducatorul
    text = str(text).strip().lower()
    replacements = {
        'ă': 'a', 'â': 'a', 'î': 'i', 'ș': 's', 'ş': 's', 'ț': 't', 'ţ': 't'
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    text = re.sub(r'\s+', ' ', text)
    return text

def scoate_sala_din_text(text):
    # Extrage materia pura. Ex: "Fizica [A05]" -> "Fizica"
    text = str(text).strip()
    text = re.sub(r'\s*\[[^\]]+\]\s*', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extrage_sala_din_text(text):
    # Extrage DOAR sala. Ex: "Fizica [A05]" -> "A05"
    text = str(text).strip()
    m = re.search(r'\[([^\]]+)\]', text)
    return m.group(1).strip() if m else ""

def inlocuieste_sala_in_text(text, sala_noua):
    # Sterge sala veche (daca exista) si o lipeste pe cea noua la final
    baza = scoate_sala_din_text(text)
    if sala_noua:
        return f"{baza} [{sala_noua}]"
    return baza

def materia_are_sala(text):
    # Verifica daca materia are deja sala atasata in vreun fel (in paranteze sau pur si simplu codul langa)
    if not text or str(text).strip() == "Liber": return False
    text = str(text)
    if re.search(r'\[[^\]]+\]', text): return True
    if re.search(r'\b[A-Za-z]{0,3}\d{2,4}[A-Za-z]?\b', text): return True
    return False

def parseaza_text_materie(text):
    # Daca vede "Fizica / Sport", stie ca Fizica e la Sapt Impara, iar Sport la Para.
    text = str(text)
    if text == 'nan' or text.strip() == '':
        return {"Impara": "Liber", "Para": "Liber"}

    if text.strip() == "Liber" or re.match(r'^-+$', text.strip()):
        return {"Impara": "Liber", "Para": "Liber"}

    linii = text.split('\n') # Taie randurile separate cu Alt+Enter
    impara_final = []
    para_final = []

    for linie in linii:
        linie = linie.strip()
        if not linie: continue

        if '/' in linie: # Daca avem materii care alterneaza
            parti = linie.split('/')
            imp = parti[0].strip() # Stanga slash-ului merge in Impara
            par = parti[1].strip() if len(parti) > 1 else "" # Dreapta merge in Para

            if imp and not re.match(r'^-+$', imp): impara_final.append(imp)
            if par and not re.match(r'^-+$', par): para_final.append(par)
        else: # Daca e materie normala saptamanala, o punem in ambele
            if not re.match(r'^-+$', linie):
                impara_final.append(linie)
                para_final.append(linie)

    val_imp = " ".join(impara_final).strip()
    val_par = " ".join(para_final).strip()

    # Returneaza un "pachet" (dictionar) cu ce facem in fiecare saptamana
    return {
        "Impara": val_imp if val_imp else "Liber",
        "Para": val_par if val_par else "Liber"
    }

def extrage_ore_slots(ora_str):
    # Transforma "10-12" din Excel intr-o lista de ore individuale: ["10-11", "11-12"]
    ora_str = str(ora_str).strip()
    # Stergem detaliile de genul "sapt 1-7" ca sa ramanem doar cu ora
    ora_str = re.sub(r'(?i)s[aă]pt[a-z\.\s]*\d{1,2}\s*-\s*\d{1,2}', '', ora_str)

    # Bug comun din Excel: Uneori "09-11" e transformat automat in data "2016-09-11". Aici il detectam.
    m_date = re.search(r'\d{4}-(\d{2})-(\d{2})', ora_str)
    if m_date:
        h1, h2 = int(m_date.group(1)), int(m_date.group(2))
        if 9 <= h1 <= 20 and 10 <= h2 <= 21 and h1 < h2:
            return [f"{h:02d}-{h+1:02d}" for h in range(h1, h2)]

    # Cautare normala pentru tiparul "OraInceput - OraSfarsit" (ex: 14-16)
    m_norm = re.search(r'([012]?\d)\s*-\s*([012]?\d)', ora_str)
    if m_norm:
        h1, h2 = int(m_norm.group(1)), int(m_norm.group(2))
        # Validam ca orele sa fie intre limitele normale (9:00 -> 21:00) ca sa evitam erori
        if 9 <= h1 <= 20 and 10 <= h2 <= 21 and h1 < h2:
            return [f"{h:02d}-{h+1:02d}" for h in range(h1, h2)]

    return []

def este_curs(text):
    t = str(text).lower()
    return "(curs)" in t or "(c)" in t

def este_activitate_sport(text):
    t = scoate_sala_din_text(text).lower().strip()
    t = re.sub(r'\s+', ' ', t)

    variante_sport = {
        "sport",
        "sport (s)",
        "sp(s)",
        "sp (s)",
        "educatie fizica",
        "educaţie fizică",
        "educatie fizica si sport 2",
        "educaţie fizică şi sport 2"
    }

    return t in variante_sport

def sala_de_sport(sala_text):
    s = str(sala_text).lower()
    return "sport" in s or "sala sport" in s


# ------------------------------------------------------------------
# 2.1 Potrivire serii / grupe (A, A-G, Toate, 432E, 3xA, 413Ba ...)
# ------------------------------------------------------------------
def seria_se_potriveste(serii_text, grupa_cautata, semigrupa=None):
    """Verifica daca textul din 'Seriile' (ex: 'A, B, 432E', 'A-G', 'Toate', '441Ab + 444Ab')
    include grupa utilizatorului. Returneaza SPECIFICITATEA potrivirii:
      0 = nu se potriveste, 1 = serie intreaga (A / A-G / Toate),
      2 = grupa anume (432E), 3 = semigrupa anume (432Eb).
    Fiind numar, se poate folosi si ca True/False. Cand pentru aceeasi disciplina exista
    atat randuri pe serie cat si pe grupa, cel cu specificitate mai mare castiga."""
    grupa = str(grupa_cautata).strip().upper()
    if len(grupa) < 4: return 0
    cifre, litera = grupa[:3], grupa[3]
    sg = semigrupa.lower() if semigrupa else None

    text = str(serii_text)
    text = re.sub(r'\([^)]*\)', ' ', text)                                   # scoate "(TST)", "(15 pers.)"
    text = re.sub(r'(?i)\b(seria|seriile|serie|grupa|grupele|grupe|series|group)\b', ' ', text)
    text = re.sub(r'(?i)\b(?:IV|III|II|I)\s*(?=[A-G]\b)', ' ', text)           # "IVC" -> "C"
    if re.search(r'(?i)\b(toate|toti|toți|all)\b', text): return 1

    best = 0
    for tok in re.split(r'[,+;/&]|\bsi\b|\band\b', text):
        tok = tok.strip().replace(' ', '')
        if not tok: continue
        m_r = re.fullmatch(r'([A-Ga-g])-([A-Ga-g])', tok)                      # interval de serii
        if m_r:
            if m_r.group(1).upper() <= litera <= m_r.group(2).upper(): best = max(best, 1)
            continue
        if re.fullmatch(r'[A-G]', tok.upper()) and len(tok) == 1:               # o singura serie
            if tok.upper() == litera: best = max(best, 1)
            continue
        m_g = re.fullmatch(r'([\dxX]{1,3})([A-G])([ab])?', tok)                # 432E, 32E, 3xA, 413Ba
        if m_g:
            num, ser, sg_tok = m_g.group(1), m_g.group(2), m_g.group(3)
            if ser != litera: continue
            if any(n.lower() != 'x' and n != cifre[-1 - i] for i, n in enumerate(reversed(num))): continue
            if sg_tok:
                if sg is None or sg_tok == sg: best = max(best, 3)
            else:
                best = max(best, 2)
    return best

# ------------------------------------------------------------------
# 2.2 Potrivire de nume (discipline) si abrevieri
# ------------------------------------------------------------------
CUVINTE_GOALE = {'de', 'si', 'in', 'a', 'ale', 'al', 'la', 'pentru', 'cu', 'din', 'pe', 'of', 'and', 'the', 'for', 'to', 'ii'}
TIPURI_ACTIVITATE = {
    'c': 'curs', 'curs': 'curs', 'course': 'curs', 'lecture': 'curs',
    'l': 'laborator', 'lab': 'laborator', 'laborator': 'laborator', 'laboratory': 'laborator',
    's': 'seminar', 'sem': 'seminar', 'seminar': 'seminar',
    'p': 'proiect', 'proiect': 'proiect', 'project': 'proiect',
}

def aliase_nume(text):
    """'Sisteme de operare \\n/ \\nCurs aplicativ...' -> ['Sisteme de operare', 'Curs aplicativ...']"""
    bucati = []
    for linie in str(text).split('\n'):
        for b in re.split(r'\s/\s|^/$|\s/$|^/\s', linie.strip()):
            b = b.strip(' /-')
            if b: bucati.append(b)
    return bucati

def baza_nume(text):
    t = normalize_key(text)
    t = re.sub(r'\([^)]*\)', ' ', t)
    t = re.sub(r'\b(curs|laborator|seminar|proiect|project|lab|sem)\b', ' ', t)
    t = re.sub(r'[^a-z0-9 ]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def _tokeni(baza):
    return [w for w in baza.split() if w not in CUVINTE_GOALE]

def _nume_simplu_potrivit(x, y):
    bx, by = baza_nume(x), baza_nume(y)
    if not bx or not by: return False
    if bx == by or difflib.SequenceMatcher(None, bx, by).ratio() >= 0.85: return True      # tolereaza greseli de scriere (Filosofie/Filozofie)
    tx, ty = _tokeni(bx), _tokeni(by)
    if not tx or not ty: return False
    def prefix(w, v): return (w == v or (len(w) >= 4 and len(v) >= 4 and (w.startswith(v) or v.startswith(w)))
                              or (len(v) >= 3 and w.endswith(v) and len(w) - len(v) <= 1) or (len(w) >= 3 and v.endswith(w) and len(v) - len(w) <= 1))
    hit = sum(1 for w in tx if any(prefix(w, v) for v in ty))
    mic = min(len(tx), len(ty))
    return hit >= min(2, mic) and hit / mic >= 0.75

def nume_potrivit(a, b):
    """True daca doua denumiri (eventual cu variante separate prin '/') se refera la aceeasi disciplina."""
    return any(_nume_simplu_potrivit(x, y) for x in aliase_nume(a) for y in aliase_nume(b))

def acronime(nume):
    """'Teoria transmisiunii informatiei' -> {'tti'};  'Semnale si sisteme 3' -> {'ss3', 'ss'};
    'Programarea Calculatoarelor ... (PCLP)' -> {'pclp', ...}"""
    rezultat = set()
    for al in aliase_nume(nume):
        for m in re.findall(r'\(([A-Za-z0-9][A-Za-z0-9\-]{1,9})\)', al):
            rezultat.add(normalize_key(m))
        cuvinte = _tokeni(baza_nume(al))
        if len(cuvinte) < 2: continue
        initiale = ''.join(w if w.isdigit() else w[0] for w in cuvinte)
        rezultat.add(initiale)
        rezultat.add(re.sub(r'\d+$', '', initiale))
    return {r for r in rezultat if len(r) >= 2}

def abreviere_din_celula(text):
    """'CIA (l) B13' -> 'cia' ; 'IEM(l)' -> 'iem' ; 'Tech. Web (p)' -> 'tech. web' ; 'Pr. CEF 2' -> ''"""
    t = scoate_sala_din_text(text)
    m = re.match(r'^\s*([^()]{1,20}?)\s*\(', t)
    return normalize_key(m.group(1)) if m else ''

def abreviere_potriveste_nume(abr, nume_norm):
    """Fallback pentru abrevieri care nu sunt initiale: 'fiz' ~ 'fizica', 'chim' ~ 'chimie', 'amp' ~ 'arhitectura microprocesoarelor'.
    Conditii: >= 3 litere, prima litera = prima litera a numelui, restul literelor apar in ordine in nume."""
    cuvinte = _tokeni(baza_nume(nume_norm))
    if len(abr) < 3 or not cuvinte or abr[0] != cuvinte[0][0]: return False
    it = iter(''.join(cuvinte))
    return all(ch in it for ch in abr)

def tip_din_celula(text):
    m = re.search(r'\(\s*([A-Za-z]+)\s*\)', str(text))
    return TIPURI_ACTIVITATE.get(m.group(1).lower(), '') if m else ''

def cheie_proiect(text):
    """'Pr. CEF 2' / 'Proiect CEF' -> 'cef' (altfel '')"""
    m = re.match(r'^\s*(?:proiect|project|pr\b\.?)\s*([A-Za-z\-]+)', scoate_sala_din_text(text), re.IGNORECASE)
    return normalize_key(m.group(1)) if m else ''

# ------------------------------------------------------------------
# 2.3 Zile / ore scrise "salbatic" (Luni 15 -17 impar, Impar Marti 9-11, Joi 13 15 par ...)
# ------------------------------------------------------------------
REGEX_ZI = re.compile(r'(?i)\b(luni|mar[tț][iî]|miercuri|joi|vineri)\b')
REGEX_ORE = re.compile(r'(?<!\d)(\d{1,2})\s*(?:-|–|—|\s)\s*(\d{1,2})(?!\d)')

def nume_zi(text):
    t = normalize_key(text)
    return {'luni': 'LUNI', 'marti': 'MARTI', 'miercuri': 'MIERCURI', 'joi': 'JOI', 'vineri': 'VINERI'}.get(t)

def _paritate_din(text):
    m = re.search(r'(?i)\b(impar|par)[ea]?\b', text)
    return None if not m else ('Impara' if m.group(1).lower() == 'impar' else 'Para')

def parseaza_sloturi(text_zi_ora, sala_text="", tip_implicit=""):
    """Transforma textul din 'Ziua si ora' in lista de sloturi:
    [{'Ziua','OraStart','OraStop','Paritate','Sala','Tip','Nota'}]. Suporta mai multe zile/ore intr-o celula,
    paritate inainte/dupa ('Impar Marti 9-11', 'Luni par 19-21', 'Miercuri 11-13 (par)') si sali pe linii."""
    linii = [l.strip() for l in str(text_zi_ora).replace('\xa0', ' ').split('\n') if l.strip()]
    sali = [s.strip() for s in str(sala_text).replace('\xa0', ' ').split('\n') if s.strip()]
    sloturi = []
    for linie in linii:
        linie_c = re.sub(r'(?i)\([^)]*s[aă]pt[^)]*\)', ' ', linie)
        linie_c = re.sub(r'(?i)s[aă]pt[a-z\.\s]*\d{1,2}\s*-\s*\d{1,2}', ' ', linie_c)
        zile = list(REGEX_ZI.finditer(linie_c))
        if not zile:
            # linie fara zi: o nota atasata ultimului slot (ex: "impar - 441B, par - 444B")
            if sloturi: sloturi[-1]['Nota'] = (sloturi[-1]['Nota'] + ' ' + linie).strip()
            continue
        for i, mz in enumerate(zile):
            sfarsit = zile[i + 1].start() if i + 1 < len(zile) else len(linie_c)
            segment = linie_c[mz.end():sfarsit]
            ziua = nume_zi(mz.group(1))
            prefix = linie_c[max(0, mz.start() - 8):mz.start()]
            m_pref = re.search(r'(?i)\b(impar|par)[ea]?\s*$', prefix)
            par_pref = _paritate_din(m_pref.group(0)) if m_pref else None
            tip = tip_din_celula(linie_c) or next((TIPURI_ACTIVITATE[w] for w in re.findall(r'(?i)\b(curs|laborator|seminar|proiect)\b', linie_c.lower())), '') or tip_implicit
            ranges = list(REGEX_ORE.finditer(segment))
            for j, mr in enumerate(ranges):
                h1, h2 = int(mr.group(1)), int(mr.group(2))
                if not (8 <= h1 < h2 <= 21): continue
                dupa = segment[mr.end(): ranges[j + 1].start() if j + 1 < len(ranges) else len(segment)]
                m_dupa = re.match(r'\s*[\(,]?\s*(impar|par)[ea]?\b', dupa, re.IGNORECASE)
                inainte = segment[: mr.start()]
                par = (_paritate_din(m_dupa.group(0)) if m_dupa else None) or _paritate_din(inainte) or par_pref or 'Ambele'
                k = len(sloturi)
                sala = sali[0] if len(sali) == 1 else (sali[k] if k < len(sali) else (sali[-1] if sali else ''))
                sloturi.append({'Ziua': ziua, 'OraStart': max(h1, 9), 'OraStop': h2, 'Paritate': par,
                                'Sala': curata_sala(sala), 'Tip': tip, 'Nota': ''})
    return sloturi

def curata_sala(text):
    """'B214 / B215' -> 'B214, B215'; '???', 'nan', 'Consultati orar' -> ''"""
    t = str(text).replace('\xa0', ' ').strip()
    if t.lower() in ('nan', '', '-', '--', '???') or 'consult' in t.lower(): return ''
    t = re.sub(r'\s*\n\s*', ', ', t)
    t = re.sub(r'\s*/\s*', ', ', t)
    return re.sub(r'\s+', ' ', t).strip(' ,')

def eticheta_slot(s):
    zi = s['Ziua'].capitalize()
    par = {'Impara': ' impar', 'Para': ' par', 'Ambele': ''}[s['Paritate']]
    sala = f" [{s['Sala']}]" if s['Sala'] else ''
    return f"{zi} {s['OraStart']}-{s['OraStop']}{par}{sala}"

def aplica_nota_paritate(slot, grupa, semigrupa):
    """Nota de tip 'impar - 441B, par - 444B' decide paritatea in functie de grupa utilizatorului."""
    nota = slot.get('Nota', '')
    gasite = re.findall(r'(?i)\b(impar|par)[ea]?\s*[-–:]\s*([0-9xX]{1,3}[A-G][ab]?)', nota)
    if not gasite: return slot
    for par, grp in gasite:
        if seria_se_potriveste(grp, grupa, semigrupa):
            slot = dict(slot); slot['Paritate'] = 'Impara' if par.lower() == 'impar' else 'Para'
            break
    return slot

# ------------------------------------------------------------------
# 2.4 Gasirea foilor (nu mai depindem de nume exacte)
# ------------------------------------------------------------------
def gaseste_foi(fisier, anul):
    """Intoarce (foaia_cu_orarul, foaia_cu_anunturi). Cauta dupa model ('AN 3', 'Anunturi - AN3', 'Anunțuri an 3'...)."""
    nume = pd.ExcelFile(fisier).sheet_names
    def n(x): return re.sub(r'\s+', '', normalize_key(x))
    foaie_orar = next((s for s in nume if re.fullmatch(rf'an0?{anul}', n(s))), None)
    foaie_anunturi = next((s for s in nume if 'anunt' in n(s) and re.search(rf'an0?{anul}(?!\d)', n(s))), None)
    return foaie_orar or f"AN {anul}", foaie_anunturi or f"Anunturi - AN{anul}"

# ==========================================
# 3. CITIRE DATE DIN EXCEL (EXTRACT & TRANSFORM)
# ==========================================
_CACHE_FOI = {}

def citeste_sheet(fisier, nume_sheet):
    # Citim fiecare foaie o singura data (e lent), apoi dam mai departe o copie ca sa nu stricam originalul
    cheie = (fisier, nume_sheet, os.path.getmtime(fisier))
    if cheie not in _CACHE_FOI:
        _CACHE_FOI[cheie] = pd.read_excel(fisier, sheet_name=nume_sheet, header=None)
    return _CACHE_FOI[cheie].copy()

def detecteaza_coloane_grupa(df, grupa_cautata):
    # Cauta capetele de tabel cu numele grupei/semigrupei (ex: 411Ca si 411Cb).
    # Daca secretariatul a scris accidental aceeasi grupa de doua ori, o folosim pe prima.
    sg_a = (grupa_cautata + "a").lower()
    sg_b = (grupa_cautata + "b").lower()
    col_sga, col_sgb = None, None

    for r in range(min(12, len(df))):
        for c in df.columns:
            val = re.sub(r'\s+', '', normalize_text(df.iloc[r, c]).lower())
            if val == sg_a and col_sga is None: col_sga = c
            elif val == sg_b and col_sgb is None: col_sgb = c

    return col_sga, col_sgb

def grupa_are_semigrupe(fisier, nume_sheet, grupa):
    # Verifica simplu daca pe coloane s-a gasit semigrupa b (pentru a cere input de la user la inceput)
    df = citeste_sheet(fisier, nume_sheet)
    col_sga, col_sgb = detecteaza_coloane_grupa(df, grupa)
    return col_sgb is not None

# ------------------------------------------------------------------
# 3.1 Foaia de ANUNTURI: parser generic de tabele
# ------------------------------------------------------------------
# Secretariatul isi schimba formatul de la un semestru la altul (coloane in alta ordine, tabele noi, alte titluri).
# De aceea NU mai citim coloanele "pe pozitie", ci recunoastem fiecare tabel dupa randul lui de antet
# (Disciplina / Seriile / Ziua si ora / Sala ...) si transformam fiecare rand intr-o "inregistrare" uniforma.

REGEX_SALA_EXACTA = re.compile(r'^[A-Za-z]{1,3}\s?\d{2,4}[A-Za-z]?$')
REGEX_SALA_IN_TEXT = re.compile(r'\b([A-Za-z]{1,3}\s?\d{2,4}[A-Za-z]?)\b')
REGEX_ANUNT_PROIECT = re.compile(
    r'(?i)seria(?:ile)?\s+([A-G](?:\s*[,+]\s*[A-G])*)\s+(\d{1,2})\s*-\s*(\d{1,2})\s+(luni|marti|miercuri|joi|vineri)\s*[,;]?\s*([A-Za-z]{1,3}\s?\d{2,4}[A-Za-z]?)')

def _rol_antet(celula):
    s = normalize_key(celula)
    if not s or len(s) > 45 or 'formular' in s or 'repartizare' in s: return None
    if 'disciplin' in s or 'subject' in s: return 'disc'
    if 'activitate' in s or s in ('tip', 'activity'): return 'activ'
    if re.search(r'\b(seria|serii|seriile|series|grupa|grupe|grupele)\b', s): return 'serii'
    if s.startswith('ziua') or s in ('ora', 'interval') or s.startswith('interval') or 'day and time' in s or s.startswith('ora '): return 'zi_ora'
    if s.startswith('sala') or s.startswith('sali') or 'room' in s: return 'sala'
    if s in ('c.d.', 'cd') or 'cadru' in s or 'profesor' in s: return 'prof'
    if 'observati' in s: return 'obs'
    return None

def _titlu_bloc(text):
    """'Compatibilitate electromagnetica\\n CEM\\n (laborator)\\nLab incepe...' -> (nume_cu_variante, tip)"""
    linii = [l.strip() for l in str(text).split('\n') if l.strip()]
    nume, tip = [], ''
    for l in linii:
        m = re.fullmatch(r'\(\s*([A-Za-z]+)\s*\)', l)
        if m and m.group(1).lower() in TIPURI_ACTIVITATE:
            tip = TIPURI_ACTIVITATE[m.group(1).lower()]
            break
        nume.append(l)
        if len(nume) >= 3: break
    if not tip: nume = nume[:1]
    return '\n'.join(nume), tip

def parseaza_anunturi(df):
    """Transforma foaia de anunturi intr-o lista de inregistrari:
       {'sect': obligatorii|optionale|facultative|alte|proiect_text, 'disc', 'activ', 'serii', 'zi_ora', 'sala', 'bloc_tip', 'in_bloc'}"""
    inreg = []
    sect, antet = 'obligatorii', {}
    bloc_nume, bloc_tip = '', ''
    ult_disc = ult_serii = ult_sala = ''
    topic_proiect, rand_topic = '', -99

    for r in range(len(df)):
        celule = ['' if pd.isna(v) else str(v).replace('\xa0', ' ').strip() for v in df.iloc[r]]
        neg = [(c, t) for c, t in enumerate(celule) if t and t.lower() != 'nan']
        if not neg: continue
        col0 = celule[0]
        col0_n = normalize_key(col0)

        # --- Anunturi in text liber despre proiecte: "Proiect CEF" + "Seria E 9-13 miercuri, B100"
        if len(neg) == 1 and neg[0][0] == 0:
            if re.match(r'^(?:proiect|project)\s+\S', col0_n) and '\n' not in col0:
                topic_proiect, rand_topic = cheie_proiect(col0), r
            for linie in col0.split('\n'):
                m = REGEX_ANUNT_PROIECT.search(normalize_key(linie).replace('marti', 'marti'))
                if m and topic_proiect and r - rand_topic <= 6:
                    inreg.append({'sect': 'proiect_text', 'disc': 'proiect ' + topic_proiect, 'topic': topic_proiect,
                                  'serii': m.group(1), 'zi_ora': f"{m.group(4)} {m.group(2)}-{m.group(3)}",
                                  'sala': m.group(5).upper().replace(' ', ''),
                                  'activ': 'proiect', 'bloc_tip': 'proiect', 'in_bloc': False})
            # titlu de sectiune?
            if len(col0) <= 60 and re.match(r'^(discipline|sali|laboratoare|alte)\b', col0_n):
                sect = _sectiune(col0_n, sect)
            continue

        # --- Rand de antet?
        roluri = {c: _rol_antet(t) for c, t in neg}
        roluri_dr = {c: ro for c, ro in roluri.items() if ro and c > 0}
        constante_sala = {c: t for c, t in neg if c > 0 and not roluri.get(c) and REGEX_SALA_EXACTA.match(t)}
        distincte = set(roluri_dr.values()) - {'obs', 'prof'}
        esteantet = len(distincte | ({'sala'} if constante_sala else set())) >= 2 or (roluri.get(0) == 'disc' and len(distincte) >= 1)
        if esteantet:
            antet = {}
            for c, ro in roluri_dr.items(): antet.setdefault(ro, c)
            if roluri.get(0) == 'disc': antet['disc'] = 0
            antet.setdefault('disc', 0)
            antet['sala_const'] = next(iter(constante_sala.values()), '')
            if col0 and (len(col0) > 60 or '\n' in col0):
                bloc_nume, bloc_tip = _titlu_bloc(col0)
                ult_disc = bloc_nume
            else:
                bloc_nume, bloc_tip = '', ''
                ult_disc = ''
            if col0 and len(col0) <= 90 and re.match(r'^(discipline|sali|laboratoare|alte)\b', col0_n):
                sect = _sectiune(col0_n, sect)
            ult_serii = ult_sala = ''
            continue

        # --- Rand de date
        if not antet: continue
        def luat(rol):
            c = antet.get(rol)
            return celule[c] if c is not None and c < len(celule) else ''
        disc_cell = luat('disc') if antet.get('disc', 0) != 0 or col0 else ''
        if disc_cell.startswith(('-', '*')) or normalize_key(disc_cell).startswith('disciplina poarta'): continue
        serii, zi_ora, sala, activ = luat('serii'), luat('zi_ora'), luat('sala'), luat('activ')
        if not sala: sala = antet.get('sala_const', '') if (antet.get('sala_const') and sect == 'alte') else ''
        if disc_cell:
            ult_disc, ult_serii, ult_sala = disc_cell, '', ''
        if not serii: serii = ult_serii
        else: ult_serii = serii
        if not sala and ult_sala and not disc_cell: sala = ult_sala      # celula de sala unita pe mai multe randuri
        elif sala: ult_sala = sala
        if not (zi_ora or sala) or not ult_disc: continue
        inreg.append({'sect': sect, 'disc': ult_disc, 'activ': normalize_key(activ), 'serii': serii,
                      'zi_ora': zi_ora, 'sala': sala, 'bloc_tip': bloc_tip, 'in_bloc': bool(bloc_nume), 'r': r})
    return inreg

def _sectiune(t, curenta):
    if 'optional' in t: return 'optionale'
    if 'facultativ' in t: return 'facultative'
    if t.startswith('alte'): return 'alte'
    if 'sali' in t or 'obligator' in t: return 'obligatorii'
    return curenta

def extrage_sali_laborator(fisier_curatat, nume_sheet_anunturi, grupa_cautata, traducator_activ, semigrupa=None):
    """Construieste dictionarul {nume/abreviere normalizata -> sala} pentru disciplinele OBLIGATORII ale grupei.
    Se salveaza fiecare disciplina sub mai multe chei: numele lung, abrevierea din orar (din traducator)
    si acronimul generat automat ('Teoria transmisiunii informatiei' -> 'tti')."""
    dataframe = citeste_sheet(fisier_curatat, nume_sheet_anunturi)
    dictionar_sali, specificitate = {}, {}

    def pune(cheie, sala, spec):
        # un rand pe grupa anume bate un rand pe intreaga serie; la egalitate, ultimul rand castiga
        if cheie and specificitate.get(cheie, 0) <= spec:
            dictionar_sali[cheie] = sala
            specificitate[cheie] = spec

    for rec in parseaza_anunturi(dataframe):
        sala = curata_sala(rec['sala'])
        if not sala: continue
        if rec['sect'] == 'proiect_text':
            spec = seria_se_potriveste(rec['serii'], grupa_cautata, semigrupa)
            if spec: pune('proiect:' + rec['topic'], sala, spec)
            continue
        optional_cu_sala = (rec['sect'] == 'optionale' and (re.search(r'(?i)laborator|proiect|\blab\b', rec['disc'])
                                                            or not parseaza_sloturi(rec['zi_ora'])))
        if rec['sect'] != 'obligatorii' and not optional_cu_sala: continue
        spec = seria_se_potriveste(rec['serii'], grupa_cautata, semigrupa)
        if not spec: continue
        if rec['sect'] != 'obligatorii':
            # optionalele: doar numele (nu acronimele generate), ca sa nu se amestece cu disciplinele obligatorii
            for alias in aliase_nume(rec['disc']):
                pune(normalize_key(re.sub(r'\s+', ' ', alias)), sala, spec)
                pune(baza_nume(alias), sala, spec)
                for abr_p in re.findall(r'\(([^)]{2,15})\)', alias):          # abrevierea scrisa in paranteza: "(SEP)", "(Tech. Web)"
                    if normalize_key(abr_p) not in TIPURI_ACTIVITATE: pune(normalize_key(abr_p), sala, spec)
            continue

        este_lab = rec['activ'] in ('', 'laborator', 'proiect') or 'lab' in rec['activ']
        for alias in aliase_nume(rec['disc']):
            disciplina_bruta = normalize_key(re.sub(r'\s+', ' ', alias))
            pune(disciplina_bruta, sala, spec)
            if este_lab:
                for ac in acronime(alias): pune(ac, sala, spec)
            if len(aliase_nume(rec['disc'])) > 1: continue   # evitam potriviri periculoase pe randuri mixte "X / Y"
            # Traducem numele lung din Anunturi in cel scurt din Orar
            for prescurtare, nume_lung in traducator_activ.items():
                p_norm, n_norm = normalize_key(prescurtare), normalize_key(nume_lung)
                if n_norm in disciplina_bruta or p_norm in disciplina_bruta:
                    # potrivirea exacta bate potrivirea prin "contine" (ca "Fizica" sa nu primeasca sala de la "Educatie fizica si sport")
                    spec_t = spec if disciplina_bruta in (n_norm, p_norm) else spec - 0.5
                    pune(p_norm, sala, spec_t)
                    pune(n_norm, sala, spec_t)
    return dictionar_sali

# ------------------------------------------------------------------
# 3.2 Foaia de ORAR: coloane + grila
# ------------------------------------------------------------------
def detecteaza_coloana_ora(df, col_sga):
    # Pleaca de la coloana cu grupa ta si merge cu -1 spre stanga pentru a gasi Timpul
    if col_sga is None: return 1
    for c in range(col_sga - 1, 0, -1):
        ora_count = 0
        for r in range(2, min(30, len(df))):
            if extrage_ore_slots(df.iloc[r, c]): ora_count += 1
        # Daca a numarat macar 3 randuri cu format logic de ora (ex 08-10), stie clar ca e coloana corecta
        if ora_count >= 3: return c
    return 1

def detecteaza_coloana_sala(df, col_sga, grupa_cautata):
    # Pleaca de la coloana ta si se uita cu +1 spre dreapta dupa Sală
    if col_sga is None: return None
    seria_mea = grupa_cautata[-1].lower()

    for c in range(col_sga + 1, len(df.columns)):
        alta_serie = False
        for r in range(min(8, len(df))):
            val = normalize_text(df.iloc[r, c]).lower()
            if val in ["sala", "sala / room", "săli"]: return c # Cauta normal dupa titlu

            # Daca da de grupa altei serii (ex: noi suntem F si vedem 431Ga), opreste scanarea spre dreapta
            if len(val) >= 4 and val[:2].isdigit() and val[3].isalpha():
                if val[3] != seria_mea: alta_serie = True

        if not alta_serie: # Cazul cand secretariatul uita sa scrie titlul Sala
            room_count = 0
            for r in range(8, min(40, len(df))):
                cell_val = normalize_text(df.iloc[r, c])
                # Verifica printr-un filtru regex daca celula arata a sala (ex: Litera + Cifre = "B125", "A05", "BN030")
                if re.match(r'^[A-Z]{1,3}\s*\d{2,4}[A-Z]?$', cell_val.strip(), re.IGNORECASE):
                    room_count += 1
            if room_count >= 2: return c # Daca gaseste minim 2, nu este o coincidenta

        if alta_serie: break # Oprim cautarea stricta

    # Doar daca 2 serii impart sala, trece de granita seriei si ia prima coloana Sală disponibila.
    for c in range(col_sga + 1, len(df.columns)):
        for r in range(min(8, len(df))):
            val = normalize_text(df.iloc[r, c]).lower()
            if val in ["sala", "sala / room", "săli"]: return c

    return None

def construieste_grila_goala(exista_sg2):
    # Deseneaza scheletul de ore (Luni 09-10, 10-11, etc.) complet liber
    grila = []
    for zi in ZILE_SAPTAMANA:
        for ora in ORE_STANDARD:
            grila.append({
                'Ziua': zi, 'Ora': ora,
                'SG1': {'Impara': 'Liber', 'Para': 'Liber'},
                'SG2': {'Impara': 'Liber' if exista_sg2 else '-', 'Para': 'Liber' if exista_sg2 else '-'},
                'SalaExcel': ''
            })
    return grila


def parseaza_alternativa(alt):
    """'Circuite Integrate Digitale 2 (curs) A103 Prof. Popescu G.' -> nume, tip, sala, profesor"""
    t = re.sub(r'\s+', ' ', str(alt)).strip(' |-/')
    t = re.sub(r'(?i)^(?:--\s*/\s*)?O\d+\.?\s*(?:\(\s*[a-z]+\s*\)\s*)?', '', t)     # marcajul "O1 (c)" al optionalelor
    m = re.match(r'^(?P<nume>.*?)\s*\(\s*(?P<tip>[A-Za-z]+)\s*\)\s*(?P<rest>.*)$', t)
    if m and m.group('tip').lower() in TIPURI_ACTIVITATE:
        nume, tip, rest = m.group('nume').strip(' -'), TIPURI_ACTIVITATE[m.group('tip').lower()], m.group('rest')
    else:
        nume, tip, rest = t, 'curs', ''
        m_s = re.match(r'^(.*?)\s*\(\s*([A-Za-z]{1,3}\s?\d{2,4}[A-Za-z]?)\s*\)\s*$', t)      # "Sociologie (A103)"
        if m_s: nume, rest = m_s.group(1), m_s.group(2)
    nume = re.sub(r'(?i)^opt\.?\s+', '', nume).strip()
    sala, prof = '', ''
    m_b = re.search(r'\[([^\]]+)\]', rest)
    if m_b:
        sala = m_b.group(1).strip()
        rest = rest.replace(m_b.group(0), ' ')
    else:
        m_s = REGEX_SALA_IN_TEXT.search(rest)
        if m_s and not re.match(r'(?i)prof', rest[m_s.start():]):
            sala = m_s.group(1).replace(' ', '')
            rest = rest.replace(m_s.group(0), ' ', 1)
    prof = re.sub(r'\s+', ' ', rest).strip(' ,;-')
    return {'Nume': nume, 'Tip': tip, 'Sala': sala, 'Prof': prof}

def agrega_pools(celule_pool):
    """Aduna orele consecutive ale aceleiasi alternative intr-un singur interval (11-12 + 12-13 -> 11-13)."""
    ore = {}
    for pc in celule_pool:
        for alt in pc['alts']:
            info = parseaza_alternativa(alt)
            ore.setdefault((pc['zi'], baza_nume(info['Nume']), info['Tip'], pc.get('par', 'Ambele')), (info, set()))[1].add(pc['ora'])
    optiuni = []
    for (zi, _, _, par), (info, ore_set) in ore.items():
        lista = sorted(ore_set)
        start = prev = lista[0]
        for h in lista[1:] + [None]:
            if h is not None and h == prev + 1:
                prev = h
                continue
            optiuni.append({**info, 'Ziua': zi, 'OraStart': start, 'OraStop': prev + 1, 'Paritate': par})
            if h is not None: start = prev = h
    return optiuni

def separa_alternative(text):
    """O celula cu mai multe materii separate prin "|" sau "||" contine ALTERNATIVE OPTIONALE (alegi una).
    Forme intalnite: "A (curs) | B (curs)",  "-- / A | B" (doar saptamani pare),  "A | B / --" (doar impare).
    Intoarce (lista_alternative, paritate); lista goala daca celula nu e de acest tip."""
    if '|' not in text: return [], 'Ambele'
    t, par = text, 'Ambele'
    m = re.match(r'^\s*(?:--|-)\s*/\s*(.+)$', t, re.S)
    if m: t, par = m.group(1), 'Para'
    else:
        m = re.match(r'^(.+?)\s*/\s*(?:--|-)\s*$', t, re.S)
        if m: t, par = m.group(1), 'Impara'
    alts = [a for a in re.split(r'\|+', t) if re.sub(r'[\s\-/]', '', a)]
    return (alts, par) if len(alts) >= 2 or '||' in text else ([], 'Ambele')

def incarca_date_in_grila(df, grila, col_ora, col_sga, col_sgb, col_sala):
    # Toarna informatia gasita pe coloanele Excelului in grila goala.
    # Daca o celula contine mai multe materii separate prin "||", ele sunt ALTERNATIVE OPTIONALE:
    # nu le punem in orarul obligatoriu, ci le intoarcem separat, ca sa le propunem in meniu.
    index_grila = {(e['Ziua'], e['Ora']): e for e in grila}
    celule_pool = []
    liber = {"Impara": "Liber", "Para": "Liber"}

    for rand in range(2, len(df)):
        ziua = normalize_text(df.iloc[rand, 0]).upper()
        if ziua not in ZILE_SAPTAMANA: continue

        ore_gasite = extrage_ore_slots(df.iloc[rand, col_ora])
        if not ore_gasite: continue

        def citeste_celula(col, implicit):
            if col is None: return implicit
            text = normalize_text(df.iloc[rand, col])
            alts, par = separa_alternative(text)
            if alts:
                for h in ore_gasite:
                    celule_pool.append({'zi': ziua, 'ora': int(h.split('-')[0]), 'alts': alts, 'par': par})
                return liber
            return parseaza_text_materie(text)

        mat_a = citeste_celula(col_sga, liber)
        mat_b = citeste_celula(col_sgb, {"Impara": "-", "Para": "-"})

        sala_excel = ""
        if col_sala is not None:
            val_s = normalize_text(df.iloc[rand, col_sala])
            if val_s.lower() not in ['nan', '', '-', '--']: sala_excel = val_s

        for ora_slot in ore_gasite:
            cheia = (ziua, ora_slot)
            if cheia not in index_grila: continue
            slot = index_grila[cheia]

            # Ataseaza materia doar daca slotul e liber
            if mat_a['Impara'] not in ["Liber", "-", ""]: slot['SG1']['Impara'] = mat_a['Impara']
            if mat_a['Para'] not in ["Liber", "-", ""]: slot['SG1']['Para'] = mat_a['Para']

            if col_sgb is not None:
                if mat_b['Impara'] not in ["Liber", "-", ""]: slot['SG2']['Impara'] = mat_b['Impara']
                if mat_b['Para'] not in ["Liber", "-", ""]: slot['SG2']['Para'] = mat_b['Para']

            if sala_excel: slot['SalaExcel'] = sala_excel

    return agrega_pools(celule_pool)

def ataseaza_sala_din_dictionar(text_materie, dictionar_sali, traducator_activ, allow_broad=True):
    # Compara o materie cu dictionarul de sali si, daca gaseste potrivire, lipeste sala cu paranteze "[A05]"
    if text_materie == "Liber": return text_materie
    if materia_are_sala(text_materie): return text_materie

    text_norm = normalize_key(text_materie)
    if text_norm in dictionar_sali:
        return f"{text_materie} [{dictionar_sali[text_norm]}]"

    # allow_broad=True permite potriviri partile (utile pentru laboratoare). La curs, o oprim.
    if allow_broad:
        # Proiecte anuntate in text liber ("Proiect CEF: Seria E, miercuri, B100") -> "Pr. CEF 2"
        cheie_pr = cheie_proiect(text_materie)
        if cheie_pr and ('proiect:' + cheie_pr) in dictionar_sali:
            return f"{text_materie} [{dictionar_sali['proiect:' + cheie_pr]}]"

        # Abrevierea din orar = acronimul numelui lung din Anunturi ("TTI (l)" = Teoria transmisiunii informatiei)
        abr = abreviere_din_celula(text_materie)
        if abr and tip_din_celula(text_materie) in ('laborator', 'proiect') and abr in dictionar_sali:
            return f"{text_materie} [{dictionar_sali[abr]}]"

        for prescurtare, nume_lung in traducator_activ.items():
            p_norm = normalize_key(prescurtare)
            n_norm = normalize_key(nume_lung)

            if p_norm in text_norm or text_norm == n_norm:
                if p_norm in dictionar_sali: return f"{text_materie} [{dictionar_sali[p_norm]}]"
                if n_norm in dictionar_sali: return f"{text_materie} [{dictionar_sali[n_norm]}]"

        # Ultima solutie: abrevierea seamana cu un singur nume lung din dictionar (Fiz -> Fizica)
        if abr and tip_din_celula(text_materie) in ('laborator', 'proiect'):
            candidati = {dictionar_sali[k] for k in dictionar_sali
                         if len(k) >= 5 and ' (l)' not in k and not k.startswith('proiect:') and abreviere_potriveste_nume(abr, k)}
            if len(candidati) == 1: return f"{text_materie} [{candidati.pop()}]"
            # greseli de scriere in orar (ex: "Phyton" in loc de "Python")
            if len(abr) >= 5:
                apropiat = difflib.get_close_matches(abr, [k for k in dictionar_sali if not k.startswith('proiect:')], n=1, cutoff=0.8)
                if apropiat: return f"{text_materie} [{dictionar_sali[apropiat[0]]}]"

    return text_materie

def ataseaza_salile_finale(grila, dictionar_sali, traducator_activ):
    # Se plimba prin toata grila si aplica functia de mai sus pentru fiecare materie
    for e in grila:
        sala_excel = e['SalaExcel']
        for sg in ['SG1', 'SG2']:
            for par in ['Impara', 'Para']:
                mat = e[sg][par]

                if not este_materie(mat):
                    if str(mat).strip() != "Liber": e[sg][par] = "Liber"
                    continue

                if este_activitate_sport(mat) and not materia_are_sala(mat):
                    if scoate_sala_din_text(mat).lower().strip() in {
                        "sport",
                        "sport (s)",
                        "sp(s)",
                        "sp (s)",
                        "educatie fizica",
                        "educaţie fizică",
                        "educatie fizica si sport 2",
                        "educaţie fizică şi sport 2"
                    }:
                        e[sg][par] = f"{mat} [Sala Sport LEU]"
                        continue

                if materia_are_sala(mat): continue

                if este_curs(mat) and not este_activitate_sport(mat):
                    mat_cu_sala = ataseaza_sala_din_dictionar(mat, dictionar_sali, traducator_activ, allow_broad=False)
                else:
                    mat_cu_sala = ataseaza_sala_din_dictionar(mat, dictionar_sali, traducator_activ, allow_broad=True)

                if mat_cu_sala != mat:
                    e[sg][par] = mat_cu_sala
                    continue

                # Daca nu a reusit sa aduca din Dictionarul de Anunturi, ataseaza sala gasita in detecteaza_coloana_sala
                if sala_excel:
                    if not este_activitate_sport(mat) and sala_de_sport(sala_excel): continue
                    e[sg][par] = f"{mat} [{sala_excel}]"


def normalizeaza_sali_cursuri_pe_zi(grila):
    # Daca sala e scrisa doar la semigrupa A (sau la alta grupa), el ia "Sala Dominanta" de pe rand si i-o da si semigrupei B.
    for zi in ZILE_SAPTAMANA:
        intrari_zi = [e for e in grila if e['Ziua'] == zi]
        for sg in ['SG1', 'SG2']:
            for par in ['Impara', 'Para']:
                sali_curs_valide = []
                for e in intrari_zi:
                    mat = e[sg][par]
                    if not este_materie(mat) or not este_curs(mat) or este_activitate_sport(mat): continue
                    sala = extrage_sala_din_text(mat)
                    if sala and not sala_de_sport(sala): sali_curs_valide.append(sala)

                if not sali_curs_valide: continue
                # Decide care e Sala Dominanta numarand cu Counter (cea mai frecventa sala pe orizontală)
                sala_dominanta = Counter(sali_curs_valide).most_common(1)[0][0]

                # Imparte sala la semigrupele care o au lipsa
                for e in intrari_zi:
                    mat = e[sg][par]
                    if not este_materie(mat) or not este_curs(mat) or este_activitate_sport(mat): continue
                    sala = extrage_sala_din_text(mat)
                    if not sala or sala_de_sport(sala):
                        e[sg][par] = inlocuieste_sala_in_text(mat, sala_dominanta)


def genereaza_orar_grupa(fisier_curatat, nume_sheet, grupa_cautata, dictionar_sali, traducator_activ):
    # Centralizeaza tot fluxul de creare a orarului pentru grupa curenta.
    # Intoarce (grila_obligatorie, optionale_gasite_in_orar)
    df = citeste_sheet(fisier_curatat, nume_sheet)
    df[0] = df[0].ffill()

    col_sga, col_sgb = detecteaza_coloane_grupa(df, grupa_cautata)
    if col_sga is None and col_sgb is None: return [], []

    col_ora = detecteaza_coloana_ora(df, col_sga if col_sga is not None else col_sgb)
    col_sala = detecteaza_coloana_sala(df, col_sga if col_sga is not None else col_sgb, grupa_cautata)

    grila = construieste_grila_goala(exista_sg2=(col_sgb is not None))
    pools = incarca_date_in_grila(df, grila, col_ora, col_sga, col_sgb, col_sala)
    ataseaza_salile_finale(grila, dictionar_sali, traducator_activ)
    normalizeaza_sali_cursuri_pe_zi(grila)

    return grila, pools

# ==========================================
# 4. EXPORT EXCEL (LOAD)
# ==========================================
def determinare_culoare(text, culori):
    # Pe baza unor cuvinte cheie in denumire (ex: "curs", "lab", "sport"), alege ce culoare de fundal primeste celula in Excel.
    text_original = str(text)
    t = scoate_sala_din_text(text_original).lower().strip()

    if "liber" in t or t == "-" or re.match(r'^-+$', t):
        return culori["liber"]

    if "⭐" in text_original or "opt " in t or "optional" in t or re.search(r'(^|\s)o[1-9][\s\.\(]', t):
        return culori["optional"]

    if "sport" in t or "educaţie fizică" in t or "educatie fizica" in t or "sp(s)" in t:
        return culori["seminar"]

    if "(curs)" in t or "(c)" in t or " curs" in t or "(course)" in t or "(lecture)" in t:
        return culori["curs"]

    if "(laborator)" in t or "(l)" in t or " lab" in t:
        return culori["laborator"]

    if "(seminar)" in t or "(s)" in t or " sem" in t:
        return culori["seminar"]

    if "(p)" in t or "proiect" in t or "pr " in t:
        return culori["proiect"]

    if not t or t == "nan":
        return None

    # Materiile necatalogate (fara etichete de "(curs)", "(lab)") primesc automat ROZ
    return culori["optional"]

def stiluri_excel():
    # Returneaza pachetul complet de setari grafice pentru OpenPyxl (Borders, Fonts, Colors)
    culori = {
        "header": PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid"),
        "curs": PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid"),
        "seminar": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
        "laborator": PatternFill(start_color="FF8080", end_color="FF8080", fill_type="solid"),
        "proiect": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        "optional": PatternFill(start_color="FF99CC", end_color="FF99CC", fill_type="solid"),
        "liber": PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    }
    stil = {
        "culori": culori,
        "font_header": Font(bold=True, color="FFFFFF", name="Calibri", size=12),
        "font_normal": Font(name="Calibri", size=11),
        "font_bold": Font(name="Calibri", size=11, bold=True),
        "aliniere": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border_subtire": Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
    }
    return stil

def aplica_stil_celula(cell, font, alignment, border, fill=None):
    # Aplicator scurt pt ca reduce numarul de linii de cod din loop-ul de desenare
    cell.font = font
    cell.alignment = alignment
    cell.border = border
    if fill: cell.fill = fill

def adauga_optionale_peste_orar(mat_sg1, mat_sg2, optionale_alese, zi, ora_start, paritate, semigrupa_utilizator, are_sg2):
    # Ia optiunile userului si suprascrie in mod tintit orarul general.
    for opt in optionale_alese:
        # Verifica daca se nimeresc la zi, ora si saptamana (impara/para)
        if opt['Ziua'] == zi and opt['OraStart'] <= ora_start < opt['OraStop']:
            if opt['Paritate'] in ["Ambele", paritate]:
                # Modifica DOAR casuta semigrupei alese, pentru a nu strica orarul de langa el
                if semigrupa_utilizator == 'a': mat_sg1 = f"⭐ {opt['Nume']}"
                elif semigrupa_utilizator == 'b' and are_sg2: mat_sg2 = f"⭐ {opt['Nume']}"
    return mat_sg1, mat_sg2

def uneste_intervale_identice(ws, start_row, end_row, are_sg2):
    # Citeste blocuri repetate de 1 ora (10-11 Fizica, 11-12 Fizica) si le transforma prin Merge Cells intr-un bloc unitar (10-12 Fizica).
    def baza(val): return scoate_sala_din_text(str(val))
    def get_sala(val): return extrage_sala_din_text(str(val))

    # Mentine in memorie o copie neafectata inainte de lipire
    valori_orig = {}
    for row in range(start_row, end_row + 1):
        valori_orig[row] = {
            'ora': str(ws.cell(row=row, column=2).value).strip(),
            'sg1': baza(ws.cell(row=row, column=3).value),
            'sg2': baza(ws.cell(row=row, column=4).value) if are_sg2 else None
        }

    # Lipește Coloana Timpului. Calculeaza ora finala din interval si le imbina DOAR daca Semigrupele fac acelasi lucru
    r = start_row
    while r <= end_row:
        b1 = valori_orig[r]['sg1']
        b2 = valori_orig[r]['sg2']
        r_end = r
        while r_end + 1 <= end_row:
            if valori_orig[r_end + 1]['sg1'] == b1 and (not are_sg2 or valori_orig[r_end + 1]['sg2'] == b2):
                r_end += 1
            else:
                break
        
        if r_end > r:
            ora_start = valori_orig[r]['ora'].split('-')[0]
            ora_stop = valori_orig[r_end]['ora'].split('-')[1]
            for i in range(r + 1, r_end + 1):
                ws.cell(row=i, column=2).value = ""
            ws.cell(row=r, column=2).value = f" {ora_start}-{ora_stop}" # Ex: 10-12
            ws.merge_cells(start_row=r, start_column=2, end_row=r_end, end_column=2)
        r = r_end + 1

    # Lipește coloanele Semigrupelor in mod separat
    coloane = [3, 4] if are_sg2 else [3]
    for col in coloane:
        col_key = 'sg1' if col == 3 else 'sg2'
        r = start_row
        while r <= end_row:
            b = valori_orig[r][col_key]
            r_end = r
            while r_end + 1 <= end_row:
                if valori_orig[r_end + 1][col_key] == b:
                    r_end += 1
                else:
                    break
            
            if r_end > r:
                valori_celule = [ws.cell(row=i, column=col).value for i in range(r, r_end + 1)]
                text_final = next((v for v in valori_celule if get_sala(v)), valori_celule[0])
                
                for i in range(r + 1, r_end + 1):
                    ws.cell(row=i, column=col).value = ""
                
                ws.cell(row=r, column=col).value = text_final
                ws.merge_cells(start_row=r, start_column=col, end_row=r_end, end_column=col)
            r = r_end + 1

def traseaza_linie_final_zi(ws, end_row, nr_coloane):
    # La final de zi (Luni, Marti) pune o bordura mai groasa la baza
    for col_idx in range(1, nr_coloane + 1):
        cell_baza = ws.cell(row=end_row, column=col_idx)
        margine_veche = cell_baza.border
        cell_baza.border = Border(left=margine_veche.left, right=margine_veche.right, top=margine_veche.top, bottom=Side(style='medium'))

def creeaza_sheet_paritate(wb, orar_baza, optionale_alese, grupa, semigrupa_utilizator, are_sg2, nume_sheet, paritate, stil):
    # Creaza o pagina (Saptamana Impara sau Para) si scrie randurile de orar in fisierul final OpenPyxl
    ws = wb.create_sheet(title=nume_sheet)
    headers = ["Ziua", "Ora", f"SG1 ({grupa}a)"]
    if are_sg2: headers.append(f"SG2 ({grupa}b)")
    ws.append(headers)

    for cell in ws[1]:
        aplica_stil_celula(cell, stil["font_header"], stil["aliniere"], stil["border_subtire"], stil["culori"]["header"])

    rand_curent = 2
    for zi in ZILE_SAPTAMANA:
        intrari_zi = [e for e in orar_baza if e['Ziua'] == zi]
        if not intrari_zi: continue

        start_row = rand_curent
        for intrare in intrari_zi:
            ora = intrare['Ora']
            ora_start = int(ora.split('-')[0])

            mat_sg1 = intrare['SG1'][paritate]
            mat_sg2 = intrare['SG2'][paritate] if are_sg2 else "-"

            # Momentul in care bagam Opționalele din consola peste grila normala (inainte sa coloram)
            mat_sg1, mat_sg2 = adauga_optionale_peste_orar(mat_sg1, mat_sg2, optionale_alese, zi, ora_start, paritate, semigrupa_utilizator, are_sg2)

            rand_nou = [zi, f" {ora}", mat_sg1]
            if are_sg2: rand_nou.append(mat_sg2)
            ws.append(rand_nou)

            # Aplica stilurile + Vopseaua potrivita
            nr_coloane = 4 if are_sg2 else 3
            for col_idx in range(1, nr_coloane + 1):
                cell = ws.cell(row=rand_curent, column=col_idx)
                aplica_stil_celula(cell, stil["font_bold"] if col_idx == 1 else stil["font_normal"], stil["aliniere"], stil["border_subtire"])

                if col_idx == 3:
                    fill = determinare_culoare(str(mat_sg1), stil["culori"])
                    if fill: cell.fill = fill
                elif col_idx == 4 and are_sg2:
                    fill = determinare_culoare(str(mat_sg2), stil["culori"])
                    if fill: cell.fill = fill

            rand_curent += 1

        end_row = rand_curent - 1
        uneste_intervale_identice(ws, start_row, end_row, are_sg2) # Lipim orele duble
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1) # Lipeste coloana de 'Luni' pe verticala

        nr_coloane = 4 if are_sg2 else 3
        traseaza_linie_final_zi(ws, end_row, nr_coloane)

    # Setam marimile coloanelor
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 45
    if are_sg2: ws.column_dimensions['D'].width = 45

def exporteaza_in_excel(orar_baza, optionale_alese, grupa, semigrupa_utilizator):
    # Sterge foaia goala initiala a Excelului, da start generarii paginilor, apoi salveaza!
    if not orar_baza: return

    are_sg2 = False
    for intrare in orar_baza:
        for par in ['Impara', 'Para']:
            if este_materie(intrare['SG2'][par]):
                are_sg2 = True
                break
        if are_sg2: break

    wb = Workbook()
    wb.remove(wb.active)
    stil = stiluri_excel()

    creeaza_sheet_paritate(wb, orar_baza, optionale_alese, grupa, semigrupa_utilizator, are_sg2, "Saptamana_Impara", "Impara", stil)
    creeaza_sheet_paritate(wb, orar_baza, optionale_alese, grupa, semigrupa_utilizator, are_sg2, "Saptamana_Para", "Para", stil)

    nume_fisier_iesire = f"Orar_Personalizat_{grupa}{semigrupa_utilizator}.xlsx"
    wb.save(nume_fisier_iesire)
    print(f"\nOrarul a fost generat si salvat ca '{nume_fisier_iesire}'!")

# ==========================================
# 5. MENIU OPTIONALE / FACULTATIVE
# ==========================================
# Sursele de informatie, in ordinea increderii:
#   1) grila de orar a grupei (alternativele separate prin "||") -> zi, ora, sala, profesor ale CURSULUI
#   2) foaia de Anunturi (tabelele de optionale / facultative / "alte anunturi") -> laboratoare, proiecte,
#      sali, plus facultativele care nu apar in grila.
# Aceeasi disciplina gasita in ambele locuri este unita intr-o singura optiune (potrivire pe nume).

NUME_TIP = {'curs': 'curs', 'laborator': 'laborator', 'seminar': 'seminar', 'proiect': 'proiect', 'activitate': 'activitate', '': 'activitate'}

def curata_nume_optiune(nume):
    # "Elemente de drept (curs)" -> "Elemente de drept" (tipul activitatii se adauga separat, o singura data)
    return re.sub(r'(?i)\s*\(\s*(curs|laborator|seminar|proiect|lecture|lab)\s*\)\s*$', '', str(nume)).strip()

def _optiune_noua(nume, tip_optiune, din_orar=False):
    nume = curata_nume_optiune(nume)
    return {'Nume': nume, 'Tip': tip_optiune, 'DinOrar': din_orar, 'Activitati': {}, 'Note': [], 'Sala': ''}

def _adauga_activitate(opt, tip, slot, spec=0, toate=False):
    tip = tip or 'activitate'
    act = opt['Activitati'].setdefault(tip, {'Tip': tip, 'Variante': [], 'Toate': toate})
    act['Variante'].append({**slot, 'Spec': spec})

def _slot_din_pool(p):
    return {'Ziua': p['Ziua'], 'OraStart': p['OraStart'], 'OraStop': p['OraStop'], 'Paritate': p['Paritate'],
            'Sala': p['Sala'], 'Prof': p['Prof'], 'Tip': p['Tip'], 'Nota': ''}

def construieste_meniu(fisier_curatat, nume_sheet_anunturi, pools, grupa, semigrupa):
    """Intoarce (lista_optiuni, lista_fara_ora). Fiecare optiune are 'Activitati' (curs / laborator / proiect ...),
    iar fiecare activitate are una sau mai multe 'Variante' (sloturi) dintre care utilizatorul alege."""
    recs = parseaza_anunturi(citeste_sheet(fisier_curatat, nume_sheet_anunturi))
    optiuni, fara_ora = [], []

    # 1) Optionalele din grila (cursul are ora exacta)
    for p in pools:
        opt = next((o for o in optiuni if o['DinOrar'] and nume_potrivit(o['Nume'], p['Nume'])), None)
        if opt is None:
            opt = _optiune_noua(p['Nume'], 'Optional', din_orar=True)
            optiuni.append(opt)
        _adauga_activitate(opt, p['Tip'], _slot_din_pool(p), spec=99, toate=True)

    # 2) Anunturi: optionale / facultative / blocuri "alte anunturi"
    for rec in recs:
        if rec['sect'] not in ('optionale', 'facultative', 'alte'): continue
        if rec['sect'] == 'alte' and not rec['in_bloc']: continue
        spec = seria_se_potriveste(rec['serii'], grupa, semigrupa) if rec['serii'] else 1
        if not spec: continue

        tip_impl = rec['bloc_tip'] or TIPURI_ACTIVITATE.get(rec['activ'], '') or tip_din_celula(rec['disc'])
        if not tip_impl and rec['sect'] == 'alte': tip_impl = 'laborator'   # blocurile din "Alte anunturi" sunt laboratoare/proiecte
        sloturi = parseaza_sloturi(rec['zi_ora'], rec['sala'], tip_impl)
        sala_rec = curata_sala(rec['sala'])
        if not sloturi and 'orar' in normalize_key(rec['zi_ora']): continue   # "conform orarului": e deja in grila, doar sala conteaza

        opt = next((o for o in optiuni if nume_potrivit(o['Nume'], rec['disc'])), None)
        if opt is None and rec['sect'] == 'alte': continue          # bloc despre o disciplina pe care n-o ai
        if opt is None:
            nume = ' / '.join(aliase_nume(rec['disc'])[:2]) or rec['disc']
            nume = re.sub(r'\s+', ' ', nume).strip()
            nume = re.sub(r'(?i)\s*\((?:curs|laborator|seminar|proiect)\)\s*$', '', nume)
            cheie = ('F' if rec['sect'] == 'facultative' else 'O', normalize_key(rec['disc']))
            opt = next((o for o in optiuni if o.get('Cheie') == cheie), None)
            if opt is None:
                opt = _optiune_noua(nume, 'Facultativ' if rec['sect'] == 'facultative' else 'Optional')
                opt['Cheie'] = cheie
                optiuni.append(opt)

        if not sloturi:
            if sala_rec and not opt['Sala']: opt['Sala'] = sala_rec
            if rec['sect'] == 'alte' and sala_rec: opt['Note'].append(f"{NUME_TIP.get(tip_impl, 'activitate')} in sala {sala_rec} (fara interval publicat)")
            continue
        for s in sloturi:
            if opt['DinOrar'] and s['Tip'] in ('curs', '') and 'curs' in opt['Activitati']:
                if s['Tip'] == 'curs' or any((v['Ziua'], v['OraStart'], v['OraStop']) == (s['Ziua'], s['OraStart'], s['OraStop'])
                                             for v in opt['Activitati']['curs']['Variante']):
                    continue                                          # cursul il stim deja exact din orar
            _adauga_activitate(opt, s['Tip'] or tip_impl, s, spec=spec)

    # 3) Completam salile cursului din tabelul "Sala Curs" si filtram variantele dupa grupa
    for opt in optiuni:
        for tip, act in list(opt['Activitati'].items()):
            if act['Toate']:
                for v in act['Variante']:
                    if not v['Sala'] and opt['Sala']: v['Sala'] = opt['Sala']
                continue
            variante = [aplica_nota_paritate(v, grupa, semigrupa) for v in act['Variante']]
            max_spec = max(v['Spec'] for v in variante)
            unice, vazute = [], set()
            for v in variante:
                cheie_v = (v['Ziua'], v['OraStart'], v['OraStop'], v['Paritate'], v['Sala'])
                if v['Spec'] == max_spec and cheie_v not in vazute:
                    vazute.add(cheie_v); unice.append(v)
            act['Variante'] = _uneste_ore_consecutive(unice)
    ramase = []
    for opt in optiuni:
        if opt['Activitati']: ramase.append(opt)
        elif opt['Tip'] == 'Optional': fara_ora.append(opt['Nume'])
    # Optionalele din Anunturi pentru seria ta care nu apar nicaieri in grila
    for rec in recs:
        if rec['sect'] == 'optionale' and not rec['zi_ora'] and (not rec['serii'] or seria_se_potriveste(rec['serii'], grupa, semigrupa)):
            nume = ' / '.join(aliase_nume(rec['disc'])[:2])
            if nume and not any(nume_potrivit(nume, o['Nume']) for o in ramase) and nume not in fara_ora:
                fara_ora.append(nume)

    def cheie_sort(o):
        s = next(iter(o['Activitati'].values()))['Variante'][0]
        return (0 if o['Tip'] == 'Optional' else 1, ZILE_SAPTAMANA.index(s['Ziua']) if s['Ziua'] in ZILE_SAPTAMANA else 9, s['OraStart'])
    ramase.sort(key=cheie_sort)
    for i, o in enumerate(ramase, 1): o['Id'] = i
    return ramase, fara_ora

def _uneste_ore_consecutive(variante):
    """Doua sloturi de cate o ora, cu aceeasi zi/sala/paritate, una dupa alta (Marti 9-10 + Marti 10-11), sunt de fapt
    o singura sedinta de doua ore."""
    if len(variante) < 2 or any(v['OraStop'] - v['OraStart'] != 1 for v in variante): return variante
    rezultat = []
    for v in sorted(variante, key=lambda x: (x['Ziua'], x['Paritate'], x['Sala'], x['OraStart'])):
        if rezultat and all(rezultat[-1][k] == v[k] for k in ('Ziua', 'Paritate', 'Sala')) and rezultat[-1]['OraStop'] == v['OraStart']:
            rezultat[-1] = {**rezultat[-1], 'OraStop': v['OraStop']}
        else:
            rezultat.append(dict(v))
    return rezultat

def afiseaza_meniu(optiuni, tip_optiune):
    for o in [x for x in optiuni if x['Tip'] == tip_optiune]:
        print(f" [{o['Id']}] {o['Nume']}")
        for act in o['Activitati'].values():
            var = act['Variante']
            if len(var) == 1 or act['Toate']:
                print(f"       {NUME_TIP[act['Tip']]}: " + ' + '.join(eticheta_slot(v) for v in var))
            else:
                print(f"       {NUME_TIP[act['Tip']]}: {len(var)} variante (alegi dupa selectare)")
        for nota in o['Note']: print(f"       ({nota})")

def _text_activitate(opt, act, v):
    prof = f" {v['Prof']}" if v.get('Prof') else ''
    sala = f" [{v['Sala']}]" if v['Sala'] else ''
    return f"{opt['Nume']} ({NUME_TIP[act['Tip']]}){prof}{sala}"

def proceseaza_alegeri_meniu(alegeri_str, dict_optiuni, optionale_alese_final):
    # Parcurge lista de numere introduse in consola (ex: "1,3") si adauga in orar activitatile fiecarei optiuni
    for numar in alegeri_str.split(','):
        numar = numar.strip()
        if not (numar.isdigit() and int(numar) in dict_optiuni): continue
        opt = dict_optiuni[int(numar)]
        print(f"\n{opt['Nume']}:")

        for act in opt['Activitati'].values():
            variante = act['Variante']
            if act['Toate'] or len(variante) == 1:
                alese = variante
                for v in alese: print(f"  S-a adaugat automat ({NUME_TIP[act['Tip']]}): {eticheta_slot(v)}")
            else:
                print(f"  Pentru {NUME_TIP[act['Tip']]} ai {len(variante)} variante:")
                for i, v in enumerate(variante, 1):
                    nota = f" - {v['Nota']}" if v.get('Nota') else ''
                    print(f"    [{i}] {eticheta_slot(v)}{nota}")
                alegere_v = input(f"  Alege varianta dorita (1-{len(variante)}, Enter = sari peste): ").strip()
                if not (alegere_v.isdigit() and 1 <= int(alegere_v) <= len(variante)):
                    print("  Sarim peste.")
                    continue
                alese = [variante[int(alegere_v) - 1]]
            for v in alese:
                optionale_alese_final.append({
                    'Nume': _text_activitate(opt, act, v), 'Ziua': v['Ziua'], 'OraStart': v['OraStart'],
                    'OraStop': v['OraStop'], 'Paritate': v['Paritate'], 'Optiune': opt['Nume']})
        for nota in opt['Note']: print(f"  Nota: {nota}")

def detecteaza_conflicte(orar_baza, optionale_alese, semigrupa_utilizator, are_sg2):
    # Avertizeaza cand o optiune aleasa se suprapune cu o materie obligatorie sau cu o alta optiune
    sg = 'SG2' if (semigrupa_utilizator == 'b' and are_sg2) else 'SG1'
    index = {(e['Ziua'], int(e['Ora'].split('-')[0])): e for e in orar_baza}
    avertismente = set()
    ocupat = {}
    for opt in optionale_alese:
        for ora in range(opt['OraStart'], opt['OraStop']):
            for par in ('Impara', 'Para'):
                if opt['Paritate'] not in ('Ambele', par): continue
                cheie = (opt['Ziua'], ora, par)
                existent = index.get((opt['Ziua'], ora))
                celula_baza = existent[sg][par] if existent else ''
                # un slot marcat "O1 (c) A | B | C" din orar este chiar locul optionalei, deci nu e conflict
                slot_optional = bool(re.search(r'(^|\s)o[1-9][\s\.\(]', normalize_key(celula_baza))) or 'optional' in normalize_key(celula_baza)
                if existent and este_materie(celula_baza) and not slot_optional and cheie not in ocupat:
                    avertismente.add(f"'{opt['Nume']}' se suprapune cu '{scoate_sala_din_text(existent[sg][par])}' ({opt['Ziua'].capitalize()} {ora}-{ora+1})")
                if cheie in ocupat and ocupat[cheie] != opt['Nume']:
                    avertismente.add(f"'{opt['Nume']}' se suprapune cu '{ocupat[cheie]}' ({opt['Ziua'].capitalize()} {ora}-{ora+1})")
                ocupat[cheie] = opt['Nume']
    return sorted(avertismente)

# ==========================================
# 6. PROGRAM PRINCIPAL
# ==========================================
def main(fisier=None):
    fisier = fisier or (sys.argv[1] if len(sys.argv) > 1 else 'orar_curatat.xlsx')

    grupa, anul = extrage_grupa_si_anul()

    # Seria G mereu engleza, Seria F engleza doar din anul 3
    este_engleza = grupa.endswith('G') or (grupa.endswith('F') and anul in ['3', '4'])
    traducator_activ = TRADUCATOR_EN if este_engleza else TRADUCATOR_RO

    try:
        pagina_an, pagina_anunturi = gaseste_foi(fisier, anul)
        are_semigrupe = grupa_are_semigrupe(fisier, pagina_an, grupa)
    except Exception as e:
        print(f"\nEroare la citirea fisierului '{fisier}' (anul {anul}).")
        print(f"   Detaliu tehnic: {e}")
        return

    # Daca exista semigrupe, il punem sa isi aleaga a sau b
    if are_semigrupe:
        semigrupa_utilizator = input("Te rog introdu semigrupa ta ('a' sau 'b'): ").strip().lower()
        if semigrupa_utilizator not in ['a', 'b']:
            print("Semigrupa invalida. Vom folosi default 'a'.")
            semigrupa_utilizator = 'a'
    else:
        semigrupa_utilizator = 'a'

    # PORNESTE FLUXUL DE PROCESARE
    try:
        sali_extrase = extrage_sali_laborator(fisier, pagina_anunturi, grupa, traducator_activ, semigrupa_utilizator)
        orar_baza, pools = genereaza_orar_grupa(fisier, pagina_an, grupa, sali_extrase, traducator_activ)
        meniu, fara_ora = construieste_meniu(fisier, pagina_anunturi, pools, grupa, semigrupa_utilizator)
    except Exception as e:
        print(f"\nEroare la citirea paginilor pentru Anul {anul}.")
        print(f"   Verifica daca fisierul contine foile '{pagina_an}' si '{pagina_anunturi}'.")
        print(f"   Detaliu tehnic: {e}")
        return

    if not orar_baza:
        print("\nNu s-a gasit orar pentru grupa specificata.")
        return

    optionale_alese_final = []

    # MENIURI OPTIONALE
    dict_optionale = {o['Id']: o for o in meniu if o['Tip'] == 'Optional'}
    if dict_optionale:
        print("\nAcestea sunt optionalele disponibile pentru grupa ta:")
        afiseaza_meniu(meniu, 'Optional')
        if fara_ora:
            print("\n(Anuntate de secretariat, dar fara interval in orarul grupei tale: " + '; '.join(fara_ora) + ")")
        alegeri_opt = input("\nScrie numerele materiilor alese (ex: 1,3) sau Enter daca nu vrei: ")
        if alegeri_opt.strip():
            proceseaza_alegeri_meniu(alegeri_opt, dict_optionale, optionale_alese_final)

    # MENIURI FACULTATIVE
    dict_facultative = {o['Id']: o for o in meniu if o['Tip'] == 'Facultativ'}
    if dict_facultative:
        rasp_fac = input("\nDoresti sa adaugi materii facultative? (Da/Nu): ").strip().lower()

        if rasp_fac == 'da':
            afiseaza_meniu(meniu, 'Facultativ')
            alegeri_fac = input("\nScrie numerele alese: ")
            if alegeri_fac.strip():
                proceseaza_alegeri_meniu(alegeri_fac, dict_facultative, optionale_alese_final)

    are_sg2_orar = any(este_materie(e['SG2'][p]) for e in orar_baza for p in ('Impara', 'Para'))
    conflicte = detecteaza_conflicte(orar_baza, optionale_alese_final, semigrupa_utilizator, are_sg2_orar)
    if conflicte:
        print("\nATENTIE - suprapuneri in orar:")
        for c in conflicte: print(f"  - {c}")

    # DATELE MERG SPRE EXCEL PENTRU EXPORTUL FINAL
    exporteaza_in_excel(orar_baza, optionale_alese_final, grupa, semigrupa_utilizator)

if __name__ == "__main__":
    main()