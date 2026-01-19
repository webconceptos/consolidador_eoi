Consolidador EOI -> Formato de Evaluación (CGR)

Este proyecto lee Expresiones de Interés (Excel/PDF) y consolida los resultados
en el formato "Formato_Salida_Expresion_Interes.xlsx", colocando a cada postulante
en un bloque de 2 columnas (F/G, H/I, J/K, ...).

Uso rápido:
1) Copia tus archivos EOI (Excel/PDF) dentro de una carpeta raíz, idealmente por proceso:
   Postulaciones/Proceso_01/*.xlsx|*.pdf
   Postulaciones/Proceso_02/*.xlsx|*.pdf
2) Edita configs/config.json (ruta de entrada, reglas de puntaje, palabras clave).
3) Ejecuta:
   python run.py

Salidas:
- outputs/Cuadro_Evaluacion.xlsx (con varios postulantes en un solo archivo, por bloques)
- outputs/consolidado.csv (tabla plana)
- outputs/log.csv (errores y advertencias)

Notas PDF:
- Si el PDF es escaneado (imagen), debes activar OCR (ver config) y tener Tesseract instalado.
