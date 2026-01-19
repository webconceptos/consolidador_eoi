
import re
from pathlib import Path

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def parse_eoi_pdf(path: Path, use_ocr: bool = False) -> dict:
    """
    Parser básico:
    - Si el PDF tiene texto, intenta extraer con pdfplumber (requiere instalar dependencia).
    - Si está escaneado y use_ocr=True, se requiere pdf2image + pytesseract.

    Para mantener el proyecto liviano, este parser devuelve un dict vacío si no hay librerías.
    """
    data = {"cursos": [], "experiencias": [], "exp_general_anios": 0.0, "exp_especifica_anios": 0.0, "java_ok": False, "oracle_ok": False}
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += "\n" + (page.extract_text() or "")
        t = text.upper()

        # Extracciones por labels comunes (ajustable)
        def after(label):
            m = re.search(rf"{re.escape(label)}\s*[:\-]?\s*(.+)", text, flags=re.IGNORECASE)
            return norm(m.group(1)) if m else ""

        data["dni"] = after("DNI")
        data["email"] = after("Email") or after("Correo")
        data["celular"] = after("Celular") or after("Teléfono")

        # nombre (heurística)
        data["nombres"] = after("Nombres")
        data["ap_paterno"] = after("Apellido paterno")
        data["ap_materno"] = after("Apellido materno")
        data["nombre_full"] = norm(f"{data.get('nombres','')} {data.get('ap_paterno','')} {data.get('ap_materno','')}")

        # cursos por keywords
        cursos = []
        for kw in ("SCRUM","RUP","ORACLE","JAVA","DESARROLLO","SOFTWARE"):
            if kw in t:
                cursos.append(kw)
        data["cursos"] = cursos

        data["java_ok"] = "JAVA" in t
        data["oracle_ok"] = "ORACLE" in t

        return data
    except Exception:
        # OCR opcional (no implementado aquí por dependencia externa)
        return data
