"""
convertir_retiros.py — convierte RETIROS-D.xlsx a data_retiros.js
para el Dashboard Ventas Emanuel (sección Finanzas Personales).

Uso: python convertir_retiros.py
Requiere: pip install openpyxl

El Excel debe tener: columna con el nombre del mes, columna "Retiro",
columna "Transf.burzaco" (o similar) y columna "Total". Encabezados en
la fila 2, datos desde la fila 3 (mismo formato que RETIROS-D.xlsx).

Si en el futuro el archivo tiene datos de más de un año, agregá una
columna "Año" antes del mes y avisame para que actualice el script.
"""

import json
import re
from pathlib import Path
import openpyxl

ARCHIVO_EXCEL = "RETIROS-D.xlsx"
ARCHIVO_SALIDA = "data_retiros.js"
ANIO = 2026  # el Excel no trae año explícito — ajustá esto si corresponde

MESES_ES = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
            "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]


def parse_num(valor):
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).replace("$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    return float(s) if s not in ("", "-", ".") else 0.0


def leer_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    registros = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        mes_txt = row[1] if len(row) > 1 else None
        if not mes_txt or not isinstance(mes_txt, str):
            continue
        mes_norm = mes_txt.strip().upper()
        if mes_norm not in MESES_ES:
            continue
        retiro = parse_num(row[2]) if len(row) > 2 else 0.0
        transf = parse_num(row[3]) if len(row) > 3 else 0.0
        total = parse_num(row[4]) if len(row) > 4 else (retiro + transf)
        registros.append({
            "a": ANIO,
            "mn": MESES_ES.index(mes_norm) + 1,
            "mes": mes_txt.strip().capitalize(),
            "retiro": round(retiro, 2),
            "transf": round(transf, 2),
            "total": round(total, 2),
        })
    registros.sort(key=lambda r: (r["a"], r["mn"]))
    return registros


def generar_js(registros, salida):
    contenido = (
        "// Generado automáticamente por convertir_retiros.py\n"
        "// No editar manualmente — modificar RETIROS-D.xlsx y volver a correr el script\n"
        f"const RAW_RETIROS = {json.dumps(registros, ensure_ascii=False, indent=2)};\n"
    )
    Path(salida).write_text(contenido, encoding="utf-8")
    total = sum(r["total"] for r in registros)
    prom = total / len(registros) if registros else 0
    print(f"✓ {len(registros)} meses exportados → {salida}")
    print(f"  Total retirado: ${total:,.0f}")
    print(f"  Promedio mensual: ${prom:,.0f}")


if __name__ == "__main__":
    if not Path(ARCHIVO_EXCEL).exists():
        print(f"✗ No se encontró '{ARCHIVO_EXCEL}' en esta carpeta.")
    else:
        registros = leer_excel(ARCHIVO_EXCEL)
        generar_js(registros, ARCHIVO_SALIDA)
        print("\n  Ahora abrí dashboard.html en tu browser 🚀")