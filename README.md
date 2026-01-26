# Consolidador EOI → Formato de Evaluación (CGR)

Este proyecto consolida **Expresiones de Interés (EOI)** en Excel/PDF y genera el
**Cuadro de Evaluación** por proceso, colocando a cada postulante en un bloque
de 2 columnas (base/score). La ejecución está orquestada por `global_ejecuta.py`,
que coordina las tareas en la carpeta `tasks/`.

## Requisitos

- Python **3.10+**
- Dependencias instaladas desde `requeriments.txt` (sí, el archivo se llama así).
- (Opcional) **Tesseract OCR** si procesarás PDFs escaneados.
- (Opcional) **OpenAI API** si vas a usar la evaluación con LLM (Task 41).

## Instalación reproducible

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requeriments.txt
```

## Estructura esperada de datos

El pipeline asume una carpeta raíz con **procesos** (p. ej. `SCI N° 068-2025`).
Cada proceso debe contener la carpeta de EOI recibidas con los postulantes
**en subcarpetas directas**:

```
<ROOT>/
  SCI N° 068-2025/
    009. EDI RECIBIDAS/
      POSTULANTE_01/
        archivo.xlsx | archivo.pdf
      POSTULANTE_02/
        archivo.xlsx | archivo.pdf
    011. INSTALACIÓN DE COMITÉ/
      Revision Preliminar*.xlsx   # plantilla base del cuadro
```

> Nota: también se acepta `009. EDI RECIBIDA`.

## Configuración

Edita `configs/config.json` para ajustar:

- `input_root` (ruta base de procesos; se usa como referencia).
- `criterios` (palabras clave y umbrales).
- `pdf.use_ocr` (activar OCR en PDFs escaneados).
- `export_calculadora` y estilos de plantilla.

Si usarás LLM:

```bash
export OPENAI_API_KEY="..."
export OPENAI_MODEL="gpt-4.1-mini"  # opcional
```

## Ejecución (pipeline completo por defecto)

El orquestador ejecuta el **pipeline mínimo**: layout → collect → init → criteria → parse → fill.

```bash
python global_ejecuta.py --root "/ruta/a/ROOT"
```

### Ejecutar pasos específicos

```bash
python global_ejecuta.py \
  --root "/ruta/a/ROOT" \
  --do-layout \
  --do-collect \
  --do-init-template \
  --do-criteria \
  --do-parse \
  --do-fill
```

### Evaluación con LLM (opcional)

```bash
python global_ejecuta.py --root "/ruta/a/ROOT" --do-eval --limit 1
```

## Salidas principales

Dentro de cada proceso (`<PROCESO>/011. INSTALACIÓN DE COMITÉ/`):

- `Cuadro_Evaluacion_<PROCESO>.xlsx` (plantilla preparada).
- `Cuadro_Evaluacion_LLENO_*.xlsx` (cuadro con postulantes).
- `files_selected.csv` / `collect_summary.json`.
- `parsed_postulantes.jsonl`.
- `criteria_evaluacion.json`.

## OCR para PDFs escaneados

Si tus PDFs son imágenes:

1) Instala Tesseract (ej.: `sudo apt-get install tesseract-ocr`).
2) Activa `pdf.use_ocr` en `configs/config.json`.

## Solución de problemas

- **No se encuentra plantilla**: asegúrate de colocar `Revision Preliminar*.xlsx`
  dentro de `011. INSTALACIÓN DE COMITÉ/` del proceso antes de ejecutar `task_15`.
- **Errores de ruta**: valida que `--root` apunte a la carpeta con los procesos,
  y que cada proceso tenga `009. EDI RECIBIDAS/`.
