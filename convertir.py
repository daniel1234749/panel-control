"""
convertir.py — convierte ventas.xlsx a data.js para el dashboard
Uso: python convertir.py
Requiere: pip install pandas openpyxl
"""

import pandas as pd
import json
from pathlib import Path

ARCHIVO_EXCEL = "ventas.xlsx"
ARCHIVO_SALIDA = "data.js"

ORDEN_MESES = [
    "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
    "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
]

def leer_excel(path):
    df = pd.read_excel(path, header=1)
    df = df.dropna(subset=["AÑO"])
    df = df.rename(columns={
        "AÑO": "anio",
        "MES": "mes",
        "TICKET": "ticket",
        "IMPORTE": "importe",
        "SUCURSALES": "sucursal"
    })
    df["anio"] = df["anio"].astype(int)
    df["ticket"] = pd.to_numeric(df["ticket"], errors="coerce").fillna(0).astype(int)
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0)
    df["mes"] = df["mes"].str.upper().str.strip()
    df["sucursal"] = df["sucursal"].str.upper().str.strip()
    df["mes_num"] = df["mes"].map({m: i+1 for i, m in enumerate(ORDEN_MESES)})
    df = df.sort_values(["anio", "mes_num", "sucursal"])
    return df

def generar_js(df, salida):
    registros = []
    for _, row in df.iterrows():
        if pd.isna(row.get("mes_num")):
            continue
        registros.append({
            "a": int(row["anio"]),
            "m": row["mes"],
            "mn": int(row["mes_num"]),
            "s": row["sucursal"],
            "t": int(row["ticket"]),
            "i": round(float(row["importe"]), 2)
        })

    analisis = {
        "total_tickets": int(df["ticket"].sum()),
        "total_importe": round(float(df["importe"].sum()), 2),
        "sucursales": sorted(df["sucursal"].unique().tolist()),
        "anos": sorted(df["anio"].unique().tolist()),
        "ticket_promedio": round(float(df["importe"].sum() / df["ticket"].sum()), 0)
    }

    contenido = (
        "// Generado automáticamente por convertir.py\n"
        "// No editar manualmente — modificar ventas.xlsx y volver a correr el script\n"
        f"const RAW = {json.dumps(registros, ensure_ascii=False, indent=2)};\n"
        f"const ANALISIS = {json.dumps(analisis, ensure_ascii=False, indent=2)};\n"
    )

    Path(salida).write_text(contenido, encoding="utf-8")
    print(f"✓ {len(registros)} registros exportados → {salida}")
    print(f"  Años: {analisis['anos']}")
    print(f"  Sucursales: {', '.join(analisis['sucursales'])}")
    print(f"  Total tickets: {analisis['total_tickets']:,}")
    print(f"  Total venta: ${analisis['total_importe']:,.0f}")
    print(f"  Ticket promedio: ${analisis['ticket_promedio']:,.0f}")

if __name__ == "__main__":
    if not Path(ARCHIVO_EXCEL).exists():
        print(f"✗ No se encontró '{ARCHIVO_EXCEL}' en esta carpeta.")
        print(f"  Asegurate de que el Excel se llama exactamente así.")
    else:
        df = leer_excel(ARCHIVO_EXCEL)
        generar_js(df, ARCHIVO_SALIDA)
        print("\n  Ahora abrí dashboard.html en tu browser 🚀")