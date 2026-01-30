# parsers/eoi_pdf.py
from __future__ import annotations

import re
from pathlib import Path
from datetime import datetime, date
import pdfplumber
import pytesseract
from typing import Dict, Any, List, Optional

DATE_RE = re.compile(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b")
EMAIL_RE = re.compile(r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[A-Za-z]{2,}\b")
DNI_RE = re.compile(r"\b(\d{8})\b")
CEL_RE = re.compile(r"(?:\+51\s*)?\b(9\d{8})\b")


def _parse_date_any(s: str) -> date | None:
    s = (s or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _norm_text(s: str) -> str:
    s = s or ""
    # Normaliza saltos y espacios para que los regex funcionen mejor
    s = s.replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()


def _after_anchor_line(text: str, anchor: str) -> str:
    """
    Devuelve el contenido de la misma línea donde aparece el anchor,
    o la siguiente línea si el anchor queda “solo”.
    """
    lines = text.splitlines()
    for i, ln in enumerate(lines):
        if anchor.lower() in ln.lower():
            # intenta mismo renglón (a la derecha)
            right = ln.lower().split(anchor.lower(), 1)[-1].strip(" :-\t")
            if right:
                return lines[i].split(anchor, 1)[-1].strip(" :-\t")
            # si quedó vacío, usa la siguiente línea si existe
            if i + 1 < len(lines):
                return lines[i + 1].strip()
    return ""


def _find_first(regex: re.Pattern, text: str) -> str:
    m = regex.search(text)
    return m.group(1) if m else ""



def _extract_name_parts(text: str, debug: bool = False, trace: list[str] | None = None) -> dict:
    import re
    def dbg(msg): 
        if trace is not None: trace.append(msg)
        if debug: print(msg)

    # recorta ventana
    m = re.search(r"I\.\s*DATOS\s+PERSONALES([\s\S]{0,2500})", text, re.IGNORECASE)
    block = (m.group(1) if m else text)
    block = re.sub(r"\s+", " ", block).strip()

    dbg("[NAME] --- BEGIN BLOCK (first 600) ---")
    dbg(block[:600])
    dbg("[NAME] --- END BLOCK ---")

    def clean(val: str) -> str:
        val = (val or "").strip(" :-\t")
        val = re.sub(r"\s+", " ", val).strip()
        return val

    # APELLIDOS: after "Apellido Materno" until next label
    apellidos = ""
    pat_ap = (
        r"Apellido\s+Materno\s*[:\-]?\s*"
        r"([A-Za-zÁÉÍÓÚÑáéíóúñ ]{3,80}?)"
        r"(?=\s+(Nombres|Lugar|Documento|Identidad|D[ií]a|Mes|Año|Celular|email|Correo)\b)"
    )
    dbg(f"[NAME] pat_ap = {pat_ap}")
    m_ap = re.search(pat_ap, block, re.IGNORECASE)
    dbg(f"[NAME] m_ap = {m_ap.group(0) if m_ap else None}")
    if m_ap:
        dbg(f"[NAME] apellidos_raw = '{m_ap.group(1)}'")
        apellidos = clean(m_ap.group(1))
    dbg(f"[NAME] apellidos = '{apellidos}'")

    # NOMBRES (intentamos 3 estrategias)
    nombres = ""

    # (1) directo: "Nombres ALEX" o "Nombres: ALEX"
    pat_nom1 = r"\bNombres\b\s*[:\-]?\s*([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,}(?:\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]{2,}){0,2})"
    dbg(f"[NAME] pat_nom1 = {pat_nom1}")
    m1 = re.search(pat_nom1, block, re.IGNORECASE)
    dbg(f"[NAME] m1 = {m1.group(0) if m1 else None}")
    if m1:
        dbg(f"[NAME] nombres_raw_1 = '{m1.group(1)}'")
        nombres = clean(m1.group(1))

    # Si "Nombres" está seguido de "Lugar de nacimiento", en este formato el nombre REAL viene
    # después de "Día Mes Año" (ver trace: "... Día Mes Año ALEX APURIMAC 30 11 1988 ...")
    if not nombres:
        m_dmy = re.search(r"\bD[ií]a\b\s+\bMes\b\s+\bA[nñ]o\b\s+([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,30})\b", block, re.IGNORECASE)
        if m_dmy:
            nombres = clean(m_dmy.group(1))


    # (2) “cola” luego de la palabra Nombres (si está mezclado con etiquetas)
    if not nombres:
        m2 = re.search(r"\bNombres\b(.{0,120})", block, re.IGNORECASE)
        dbg(f"[NAME] m2_tail = '{m2.group(1) if m2 else None}'")
        if m2:
            tail = m2.group(1)
            # toma primera palabra que no sea etiqueta
            pat_word = r"\b(?!Lugar\b|Documento\b|Apellido\b|Identidad\b|Celular\b|email\b|Correo\b)([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,})\b"
            mword = re.search(pat_word, tail, re.IGNORECASE)
            dbg(f"[NAME] mword = {mword.group(0) if mword else None}")
            if mword:
                nombres = clean(mword.group(1))

    # (3) fallback: si no existe “Nombres” en el block, buscar un “ALEX” tipo nombre cerca del DNI
    if not nombres:
        dbg("[NAME] Fallback 3: buscar nombre cerca del DNI")
        dni_m = re.search(r"\b(\d{8})\b", block)
        dbg(f"[NAME] dni_in_block = {dni_m.group(1) if dni_m else None}")
        if dni_m:
            start = max(0, dni_m.start() - 80)
            end = min(len(block), dni_m.end() + 220)
            window = block[start:end]
            dbg("[NAME] window_after_dni = " + window)
            # buscar una palabra que no sea etiqueta, típica de nombre (2-20 chars)
            mname = re.search(r"\b(?!Peruano\b|Peruana\b|Lugar\b|Apellido\b|Documento\b|Identidad\b)([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,20})\b", window, re.IGNORECASE)
            dbg(f"[NAME] mname = {mname.group(1) if mname else None}")
            if mname:
                nombres = clean(mname.group(1))

    # saneo
# Si el parser capturó "Lugar de nacimiento" como nombres (caso típico de tabla),
# entonces el nombre real viene después de "Día Mes Año"
    if not nombres or nombres.lower() in ("lugar", "lugar de nacimiento"):
        nombres = ""
        m_dmy = re.search(
            r"\bD[ií]a\b\s+\bMes\b\s+\bA[nñ]o\b\s+([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,30})\b",
            block,
            re.IGNORECASE,
        )
        if m_dmy:
            nombres = clean(m_dmy.group(1))

    dbg(f"[NAME] nombres = '{nombres}'")

    nombre_full = clean(f"{apellidos} {nombres}")
    dbg(f"[NAME] nombre_full = '{nombre_full}'")

    return {
        "apellidos": apellidos,
        "nombres": nombres,
        "nombre_full": nombre_full,
    }

def _extract_name_parts_old(text: str) -> dict:
    """
    Extrae apellidos y nombres desde el bloque 'I. DATOS PERSONALES' sin depender de MAYÚSCULAS.
    Corta la captura cuando detecta que empieza otra etiqueta (Nombres/Lugar/Celular/email/etc.).
    """
    import re

    # Recorta ventana de datos personales para no contaminar con el resto del CV
    m = re.search(r"I\.\s*DATOS\s+PERSONALES([\s\S]{0,2500})", text, re.IGNORECASE)
    block = (m.group(1) if m else text)
    block = re.sub(r"\s+", " ", block).strip()

    # Helper: limpia basura típica de etiquetas que a veces se mete
    def clean(val: str) -> str:
        val = (val or "").strip(" :-\t")
        val = re.sub(r"\s+", " ", val).strip()
        return val

    # --- APELLIDOS ---
    # Caso típico: "Apellido Paterno Apellido Materno MANSILLA ZUÑIGA Nombres ALEX ..."
    # Captura después de "Apellido Materno" hasta antes de la siguiente etiqueta
    apellidos = ""
    m_ap = re.search(
        r"Apellido\s+Materno\s*[:\-]?\s*"
        r"([A-Za-zÁÉÍÓÚÑáéíóúñ ]{3,80}?)"
        r"(?=\s+(Nombres|Lugar|Documento|Identidad|D[ií]a|Mes|Año|Celular|email|Correo)\b)",
        block,
        re.IGNORECASE,
    )
    if m_ap:
        apellidos = clean(m_ap.group(1))

    # --- NOMBRES ---
    nombres = ""
    # Intento 1: "Nombres: ALEX" o "Nombres ALEX"
    m_nom = re.search(
        r"\bNombres\b\s*[:\-]?\s*([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,}(?:\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]{2,}){0,2})",
        block,
        re.IGNORECASE,
    )
    if m_nom:
        nombres = clean(m_nom.group(1))

    # Intento 2: si el PDF “rompe” la celda, a veces queda "Nombres" suelto y el valor está pegado a otra etiqueta.
    # Buscamos la primera palabra tipo nombre que aparezca DESPUÉS de la palabra "Nombres"
    if not nombres:
        m2 = re.search(r"\bNombres\b(.{0,80})", block, re.IGNORECASE)
        if m2:
            tail = m2.group(1)
            # saca la primera palabra "decente" que no sea etiqueta
            mword = re.search(r"\b(?!Lugar\b|Documento\b|Apellido\b|Identidad\b|Celular\b|email|Correo\b)([A-Za-zÁÉÍÓÚÑáéíóúñ]{2,})\b", tail, re.IGNORECASE)
            if mword:
                nombres = clean(mword.group(1))

    # Hardening: si “nombres” quedó como “Lugar” o vacío, descártalo
    if nombres.lower() in ("lugar", "lugar de nacimiento"):
        nombres = ""

    # A veces “Apellido Materno” captura cosas extra; quédate con 2-3 palabras si hay mucho ruido
    if apellidos:
        # Solo letras/espacios, corta si se coló algo raro
        apellidos = re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ ]+", " ", apellidos)
        apellidos = clean(apellidos)

    nombre_full = clean(f"{apellidos} {nombres}")

    return {
        "apellidos": apellidos,
        "nombres": nombres,
        "nombre_full": nombre_full,
    }

###### FORMACIÓN ACADÉMICA ############
#
###

def _norm_text(s: str) -> str:
    # asumo que ya tienes una; dejo una segura si no:
    s = re.sub(r"[ \t]+", " ", s or "").strip()
    s = re.sub(r"\s*\n\s*", "\n", s)
    return s.strip()

def _extract_section(text: str, start_pat: str, end_pats: List[str], max_len: int = 4000) -> str:
    """Extrae una sección por encabezado y la corta antes del siguiente encabezado."""
    m = re.search(start_pat, text, re.IGNORECASE)
    if not m:
        return ""
    chunk = text[m.end(): m.end() + max_len]
    # cortar por el primer end_pat que aparezca
    cut = len(chunk)
    for ep in end_pats:
        me = re.search(ep, chunk, re.IGNORECASE)
        if me:
            cut = min(cut, me.start())
    return _norm_text(chunk[:cut])

def _extract_education(text: str) -> Dict[str, Any]:
    """
    Heurística para CV "Formato Contraloría/BID":
    - Busca el bloque 'FORMACION ACADEMICA'
    - Extrae entradas tipo: TITULO/BACHILLER/EGRESADO/... + carrera + fecha + ciudad/pais
    - Asocia universidad cercana (en líneas contiguas).
    """
    text_n = _norm_text(text)

    # 1) Aislar sección (evita ruido de otras partes)
    sec = _extract_section(
        text_n,
        start_pat=r"\bFORMACION\s+ACADEMICA\b",
        end_pats=[
            r"\bb\.?1\)",            # b.1) cursos
            r"\bCURSO\b",            # cursos
            r"\bESTUDIOS\s+COMPLEMENTARIOS\b",
            r"\bEXPERIENCIA\b",
        ],
    )

    # fallback si no hay encabezado exacto
    if not sec:
        sec = text_n

    lines = [ln.strip() for ln in sec.splitlines() if ln.strip()]

    # 2) Diccionario de salida (permitimos múltiples grados)
    out: Dict[str, Any] = {
        "items": [],   # lista de formaciones detectadas
        "flags": {     # presencia general
            "bachiller": False,
            "egresado": False,
            "titulo": False,
            "maestria": False,
            "egresado_maestria": False,
            "colegiatura": False,
        }
    }

    # 3) Flags globales (por si el formato solo lista palabras)
    def flag(pat: str) -> bool:
        return bool(re.search(pat, sec, re.IGNORECASE))

    out["flags"]["bachiller"] = flag(r"\bBACHILLER\b")
    out["flags"]["egresado"] = flag(r"\bEGRESADO\b")
    out["flags"]["titulo"] = flag(r"\bTITULO\b")
    out["flags"]["maestria"] = flag(r"\bMAESTR(I|Í)A\b")
    out["flags"]["egresado_maestria"] = flag(r"\bEGRESADO\s+DE\s+MAESTR(I|Í)A\b")
    out["flags"]["colegiatura"] = flag(r"\bCOLEGIATURA\b")

    # 4) Regex para capturar filas tipo tabla:
    #    (grado) (carrera/especialidad) (fecha dd/mm/yyyy o d/m/yyyy) (ciudad/pais)
    row_re = re.compile(
        r"\b(?P<grado>TITULO|TÍTULO|BACHILLER|EGRESADO(?:\s+UNIVERSITARIO)?|MAESTR(I|Í)A|EGRESADO\s+DE\s+MAESTR(I|Í)A)\b"
        r"\s+(?P<carrera>[A-ZÁÉÍÓÚÑa-záéíóúñ./()\- ]{3,80}?)"
        r"\s+(?P<fecha>\d{1,2}/\d{1,2}/\d{2,4})"
        r"\s+(?P<lugar>[A-ZÁÉÍÓÚÑa-záéíóúñ./\- ]{3,40})\b",
        re.IGNORECASE
    )

    # 5) Helper: detectar universidad en una línea (y unir si está partida)
    def is_uni_line(ln: str) -> bool:
        return bool(re.search(r"\bUNIVERSIDAD\b", ln, re.IGNORECASE))

    # índice → línea universidad “compuesta” (si está partida en 2-3 líneas)
    uni_lines: Dict[int, str] = {}
    i = 0
    while i < len(lines):
        if is_uni_line(lines[i]):
            uni = lines[i]
            j = i + 1
            # unir líneas contiguas que parecen continuar el nombre
            while j < len(lines) and not row_re.search(lines[j]) and not re.match(r"^[ab]\)?", lines[j], re.I):
                # corta si aparece otra cabecera fuerte
                if re.search(r"\b(BACHILLER|TITULO|EGRESADO|MAESTR(I|Í)A|COLEGIATURA)\b", lines[j], re.I):
                    break
                # si es claramente otra cosa (fechas, encabezados), paramos
                if re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", lines[j]):
                    break
                if len(lines[j]) <= 3:
                    break
                # une
                uni += " " + lines[j]
                j += 1
            uni_lines[i] = _norm_text(uni.replace("\n", " "))
            i = j
        else:
            i += 1

    # 6) Parse de filas + asignación de universidad “cercana”
    for idx, ln in enumerate(lines):
        m = row_re.search(ln)
        if not m:
            continue

        grado = _norm_text(m.group("grado").upper().replace("TÍTULO", "TITULO"))
        carrera = _norm_text(m.group("carrera"))
        fecha = _norm_text(m.group("fecha"))
        lugar = _norm_text(m.group("lugar"))

        # buscar universidad más cercana (unas líneas arriba/abajo)
        uni: Optional[str] = None
        window = range(max(0, idx - 5), min(len(lines), idx + 6))
        # prioridad: línea con UNIVERSIDAD más cercana
        best_dist = 999
        for j in window:
            if j in uni_lines:
                dist = abs(j - idx)
                if dist < best_dist:
                    best_dist = dist
                    uni = uni_lines[j]

        out["items"].append({
            "grado": grado,
            "carrera": carrera,
            "fecha": fecha,
            "lugar": lugar,
            "universidad": uni or "",
            "line": idx + 1,
            "raw": ln,
        })

    return out
##
#
#
####### FIN DE FORMACIÓN ACADÉMICA



def _extract_contact(text: str) -> dict:
    # Preferir DNI por regex (8 dígitos) porque el PDF viene "tabulado"
    dni = _find_first(DNI_RE, text)

    celular = _find_first(CEL_RE, text)
    email_m = EMAIL_RE.search(text)
    email = email_m.group(0) if email_m else ""

    return {"dni": dni, "celular": celular, "email": email}

def _split_apellidos(apellidos: str) -> tuple[str, str]:
    parts = [p for p in (apellidos or "").split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])

def _is_scanned_pdf(page_texts: list[str]) -> bool:
    if not page_texts:
        return False
    total_chars = sum(len((t or "").strip()) for t in page_texts)
    avg_chars = total_chars / max(len(page_texts), 1)
    non_empty_pages = sum(1 for t in page_texts if len((t or "").strip()) >= 10)
    empty_ratio = 1 - (non_empty_pages / max(len(page_texts), 1))
    return avg_chars < 60 or empty_ratio >= 0.6

def _extract_pdf_text(pdf_path: Path, use_ocr: bool, debug: bool, trace: list[str]) -> tuple[str, bool, bool]:
    page_texts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in pdf.pages:
            page_texts.append(pg.extract_text() or "")

        is_scanned = _is_scanned_pdf(page_texts)
        print("Es escaneado ?")
        print(is_scanned)

        print("Es use_ocr ?")
        print(use_ocr)


        ocr_used = use_ocr
        if is_scanned and use_ocr:
            ocr_used = True
            page_texts = []
            for idx, pg in enumerate(pdf.pages, start=1):
                _dbg(trace, f"[PDF] OCR page {idx}", debug)
                img = pg.to_image(resolution=300).original
                ocr_text = pytesseract.image_to_string(img, lang="spa") or ""
                page_texts.append(ocr_text)
                
    return "\n".join(page_texts), is_scanned, ocr_used

def _build_formacion_obligatoria(edu: dict) -> dict:
    items = []
    resumen_parts = []

    print("Educación:")
    print(edu)

    uni = (edu.get("universidad") or "").strip()
    for key in ("bachiller", "egresado", "titulo"):
        label = (edu.get(key) or "").strip()
        if not label:
            continue
        item = {
            "titulo_item": label,
            "especialidad": "",
            "fecha": "",
            "centro": uni,
            "ciudad": "",
        }
        items.append(item)
        if uni:
            resumen_parts.append(f"{label}: ({uni})")
        else:
            resumen_parts.append(label)

    return {
        "items": items,
        "resumen": " ; ".join(resumen_parts).strip(),
        "meta": {"source": "pdf"},
    }

def _build_estudios_complementarios(cursos: list[str]) -> dict:
    blocks = []
    total_horas = 0
    resumen = ""
    if cursos:
        items = [
            {
                "nro": str(idx),
                "centro": "",
                "capacitacion": curso,
                "fecha_inicio": "",
                "fecha_fin": "",
                "horas": 0,
            }
            for idx, curso in enumerate(cursos, start=1)
        ]
        resumen = "\n".join(cursos).strip()
        blocks.append(
            {
                "id": "b.1",
                "row": None,
                "title": "ESTUDIOS COMPLEMENTARIOS",
                "items": items,
                "total_horas": total_horas,
                "resumen": resumen,
            }
        )

    return {
        "blocks": blocks,
        "total_horas": total_horas,
        "resumen": resumen,
    }

def _days_between(d1: date | None, d2: date | None) -> int:
    if not d1 or not d2:
        return 0
    if d2 < d1:
        return 0
    return int((d2 - d1).days) + 1

def _build_experiencia_block(pairs: list[tuple[str, str]], label: str) -> dict:
    items = []
    resumen_lines = []
    total_dias = 0
    for idx, (fi, ff) in enumerate(pairs, start=1):
        d1 = _parse_date_any(fi)
        d2 = _parse_date_any(ff)
        dias = _days_between(d1, d2)
        total_dias += dias
        item = {
            "row": idx,
            "nro": str(idx),
            "entidad": "",
            "proyecto": "",
            "cargo": "",
            "fecha_inicio": fi,
            "fecha_fin": ff,
            "tiempo_en_cargo": "",
            "dias_calc": dias,
            "descripcion": "",
        }
        items.append(item)
        line = " | ".join([p for p in [fi, ff] if p]).strip()
        if line:
            resumen_lines.append(line)

    return {
        "items": items,
        "total_dias_calc": total_dias,
        "resumen": "\n\n".join(resumen_lines).strip(),
        "_meta": {"source": "pdf", "label": label},
    }

def _slice_section(text: str, start_anchor: str, end_anchor: str | None) -> str:
    t_low = text.lower()
    s = t_low.find(start_anchor.lower())
    if s < 0:
        return ""
    if end_anchor:
        e = t_low.find(end_anchor.lower(), s + 1)
        if e > s:
            return text[s:e]
    return text[s:]

def _extract_date_pairs(section_text: str) -> list[tuple[str, str]]:
    dates = DATE_RE.findall(section_text)
    # Empareja de dos en dos (inicio, fin) en orden de aparición
    pairs = []
    for i in range(0, len(dates) - 1, 2):
        fi = dates[i]
        ff = dates[i + 1]
        pairs.append((fi, ff))
    return pairs

def _dbg(out_lines: list[str], msg: str, debug: bool):
    out_lines.append(msg)
    if debug:
        print(msg)

def parse_eoi_pdf_pro(pdf_path: Path, use_ocr: bool = False, debug: bool = False) -> dict:

    trace = []
    _dbg(trace, f"[PDF] file = {pdf_path}", debug)

    pdf_path = Path(pdf_path)

    # --- Extrae texto ---
    raw, is_scanned, ocr_used = _extract_pdf_text(pdf_path, use_ocr, debug, trace)
    text = _norm_text(raw)
    print("Texto PDF normalizado")
    print(text)
    # --- Debug forense ---
    debug_dir = pdf_path.parent / "_debug_pdfs"
    debug_dir.mkdir(parents=True, exist_ok=True)
    (debug_dir / f"{pdf_path.stem}.txt").write_text(text, encoding="utf-8")

    # Si el texto es muy corto y no usamos OCR, señalizamos.
    if is_scanned and not use_ocr:
        return {"needs_ocr": True, "is_scanned": True, "source_file": str(pdf_path)}

    # --- Campos base ---
    #name_parts = _extract_name_parts(text)
    name_parts = _extract_name_parts(text, debug=debug, trace=trace)
    print("Nombre de las partes")
    print(name_parts)

    contact = _extract_contact(text)
    edu = _extract_education(text)
    apellido_paterno, apellido_materno = _split_apellidos(name_parts.get("apellidos", ""))

    formacion_obligatoria = _build_formacion_obligatoria(edu)

    # --- Cursos: capturamos líneas cercanas a PLATZI/UDEMY/ISO/ENFAE como lista ---
    cursos = []
    for ln in text.splitlines():
        u = ln.upper()
        if any(k in u for k in ("PLATZI", "UDEMY", "ISO/IEC", "ISO", "ENFAE", "ARGOS", "KUNAK", "NEW HORIZONTS")):
            ln2 = ln.strip()
            if len(ln2) >= 6:
                cursos.append(ln2)
    # elimina duplicados preservando orden
    seen = set()
    cursos_uniq = []
    for c in cursos:
        key = c.lower()
        if key not in seen:
            seen.add(key)
            cursos_uniq.append(c)

    # --- Experiencia: extraer intervalos desde secciones ---
    sec_gen = _slice_section(text, "a) EXPERIENCIA GENERAL", "b) EXPERIENCIA ESPECIFICA 1")
    sec_esp = _slice_section(text, "b) EXPERIENCIA ESPECIFICA 1", "b) EXPERIENCIA ESPECIFICA 2")
    # Si en algunos PDFs los anchors varían, agrega más fallbacks aquí.

    gen_pairs = _extract_date_pairs(sec_gen)
    esp_pairs = _extract_date_pairs(sec_esp)

    # Convierte a estructura uniforme
    exp_general = _build_experiencia_block(gen_pairs, "general")
    exp_especifica = _build_experiencia_block(esp_pairs, "especifica")
    gen_days = int(exp_general.get("total_dias_calc") or 0)
    esp_days = int(exp_especifica.get("total_dias_calc") or 0)

    # Señales deseables por texto
    t_up = text.upper()
    java_ok = " JAVA" in t_up or "SPRING" in t_up or "SPRING BOOT" in t_up
    oracle_ok = "ORACLE" in t_up or "PL/SQL" in t_up or "PL-SQL" in t_up

    debug_dir = pdf_path.parent / "_debug_pdfs"
    debug_dir.mkdir(parents=True, exist_ok=True)
    (debug_dir / f"{pdf_path.stem}__trace.txt").write_text("\n".join(trace), encoding="utf-8")


    out = {
        "source_file": str(pdf_path),
        "needs_ocr": False,
        "is_scanned": bool(is_scanned),
        "ocr_used": bool(ocr_used),
        **name_parts,
        "apellido_paterno": apellido_paterno,
        "apellido_materno": apellido_materno,
        **contact,
        "formacion_obligatoria": formacion_obligatoria,
        "estudios_complementarios": _build_estudios_complementarios(cursos_uniq),
        "cursos": cursos_uniq,
        "exp_general": exp_general,
        "exp_especifica": exp_especifica,
        "exp_general_dias": gen_days,
        "exp_especifica_dias": esp_days,
        "java_ok": bool(java_ok),
        "oracle_ok": bool(oracle_ok),
    }

    out["exp_general_resumen_text"] = (exp_general.get("resumen") or "").strip()
    out["exp_especifica_resumen_text"] = (exp_especifica.get("resumen") or "").strip()

    def to_ymd(dias: int) -> str:
        if dias <= 0:
            return "0 año(s), 0 mes(es), 0 día(s)"
        anios = dias // 365
        rem = dias % 365
        meses = rem // 30
        dd = rem % 30
        return f"{anios} año(s), {meses} mes(es), {dd} día(s)"

    out["exp_general_total_text"] = to_ymd(out["exp_general_dias"])
    out["exp_especifica_total_text"] = to_ymd(out["exp_especifica_dias"])

    return out
