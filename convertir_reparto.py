"""
convertir_reparto.py — convierte el export de Don Gestión a data_reparto.js
para el Dashboard Ventas Emanuel (sección Reparto Tucumán), y muestra el
total vendido por mes.

Uso: python convertir_reparto.py
Requiere: pip install openpyxl

v3:
- Lee fechas tanto si vienen como fecha de Excel como si vienen como texto
  ("31/08/2026 11:23"). La v2 descartaba sin avisar las filas con fecha en
  texto (en el export de 2026 eran todas las de agosto y septiembre).
- Imprime una tabla de totales por mes (ventas, total, cobrado, no cobrado).
- Avisa si quedó afuera alguna fila que parecía una venta.
"""

import datetime as dt
import json
import re
from collections import defaultdict
from pathlib import Path
from statistics import median

import openpyxl

CANDIDATOS_EXCEL = [
    "Todas las ventas - Mayorista Emanuel.xlsx",
    "Todas_las_ventas_-_Mayorista_Emanuel.xlsx",
    "reparto_emanuel.xlsx",
    "reparto.xlsx",
]
ARCHIVO_SALIDA = "data_reparto.js"

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

FORMATOS_FECHA = ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d")


def parse_monto(valor):
    """Convierte '$ 332.200,96' -> 332200.96. Soporta números ya numéricos."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).replace("$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    return float(s) if s not in ("", "-", ".") else 0.0


def parse_fecha(valor):
    """Acepta fecha de Excel (datetime) o texto tipo '31/08/2026 11:23'.
    Devuelve None si no es una fecha (títulos, encabezados, 'Total', etc.)."""
    if isinstance(valor, dt.datetime):
        return valor
    if isinstance(valor, dt.date):
        return dt.datetime(valor.year, valor.month, valor.day)
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in FORMATOS_FECHA:
            try:
                return dt.datetime.strptime(texto, formato)
            except ValueError:
                pass
    return None


def pesos(n):
    """Formato argentino: 1234567.8 -> '$ 1.234.567,80'."""
    return "$ " + f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def encontrar_archivo():
    for nombre in CANDIDATOS_EXCEL:
        if Path(nombre).exists():
            return nombre
    return None


def leer_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    registros = []
    vistos = {}
    descartadas = 0
    sospechosas = []

    for nro_fila, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        fecha_cruda = row[0] if len(row) > 0 else None
        cliente = row[1] if len(row) > 1 else None
        estado = row[4] if len(row) > 4 else None
        total = row[5] if len(row) > 5 else None
        nro_factura = row[11] if len(row) > 11 else None

        fecha = parse_fecha(fecha_cruda)

        # Una venta real tiene fecha, cliente y número de factura.
        # Títulos, encabezados repetidos, filas "Total" y botones se descartan.
        if fecha is None or nro_factura is None or cliente is None:
            descartadas += 1
            # Si tiene factura numérica y monto pero no se pudo leer la fecha,
            # es una venta que se perdería: avisar.
            if isinstance(nro_factura, (int, float)) and total is not None:
                sospechosas.append((nro_fila, fecha_cruda, cliente, total))
            continue

        registro = {
            "f": fecha.strftime("%Y-%m-%d"),
            "a": fecha.year,
            "mn": fecha.month,
            "c": str(cliente).strip(),
            "e": str(estado).strip() if estado else "Sin estado",
            "t": round(parse_monto(total), 2),
            "_nf": nro_factura,
        }

        if nro_factura in vistos:
            anterior = vistos[nro_factura]
            if (anterior["f"] != registro["f"] or anterior["c"] != registro["c"]
                    or abs(anterior["t"] - registro["t"]) > 0.5):
                print(f"  ⚠ Factura {nro_factura} aparece más de una vez con datos distintos — se usó la primera encontrada.")
            continue

        vistos[nro_factura] = registro
        registros.append(registro)

    print(f"  ({descartadas} filas descartadas: títulos, encabezados, subtotales o filas vacías)")
    if sospechosas:
        print("\n⚠ Estas filas parecen ventas pero no se pudo leer la fecha (quedaron AFUERA):")
        for fila, fecha, cliente, total in sospechosas:
            print(f"   Fila {fila}: fecha={fecha!r}  cliente={cliente}  total={total}")

    revisar_facturas_faltantes(list(vistos.keys()))

    for r in registros:
        del r["_nf"]
    registros.sort(key=lambda r: r["f"])
    return registros


def revisar_facturas_faltantes(numeros):
    nums = sorted(int(n) for n in numeros if isinstance(n, (int, float)))
    if not nums:
        return
    faltantes = sorted(set(range(nums[0], nums[-1] + 1)) - set(nums))
    if faltantes:
        print(f"\n  ℹ Números de factura que no están en el archivo ({len(faltantes)}): "
              f"{', '.join(str(n) for n in faltantes[:30])}{' ...' if len(faltantes) > 30 else ''}")
        print("    (pueden ser anuladas o de otro canal; conviene confirmarlo en Don Gestión)")


def totales_por_mes(registros):
    resumen = defaultdict(lambda: {"ventas": 0, "total": 0.0, "cobrado": 0.0, "no_cobrado": 0.0})
    for r in registros:
        m = resumen[(r["a"], r["mn"])]
        m["ventas"] += 1
        m["total"] += r["t"]
        if r["e"].lower() == "cobrado":
            m["cobrado"] += r["t"]
        else:
            m["no_cobrado"] += r["t"]

    linea = "-" * 84
    print("\nVENTAS POR MES")
    print(linea)
    print(f"{'Mes':<17}{'Ventas':>7}{'Total vendido':>21}{'Cobrado':>20}{'No cobrado':>19}")
    print(linea)
    g = {"ventas": 0, "total": 0.0, "cobrado": 0.0, "no_cobrado": 0.0}
    for (anio, mes) in sorted(resumen):
        m = resumen[(anio, mes)]
        nombre = f"{MESES[mes]} {anio}"
        print(f"{nombre:<17}{m['ventas']:>7}{pesos(m['total']):>21}{pesos(m['cobrado']):>20}{pesos(m['no_cobrado']):>19}")
        for k in g:
            g[k] += m[k]
    print(linea)
    print(f"{'TOTAL':<17}{g['ventas']:>7}{pesos(g['total']):>21}{pesos(g['cobrado']):>20}{pesos(g['no_cobrado']):>19}")
    print(linea)

    # También queda guardado en un CSV que se abre con Excel.
    with open("totales_por_mes.csv", "w", encoding="utf-8-sig") as f:
        f.write("Mes;Ventas;Total vendido;Cobrado;No cobrado\n")
        for (anio, mes) in sorted(resumen):
            m = resumen[(anio, mes)]
            f.write(f"{MESES[mes]} {anio};{m['ventas']};{m['total']:.2f};{m['cobrado']:.2f};{m['no_cobrado']:.2f}\n".replace(".", ","))
        f.write(f"TOTAL;{g['ventas']};{g['total']:.2f};{g['cobrado']:.2f};{g['no_cobrado']:.2f}\n".replace(".", ","))
    print("  ✓ También guardado en totales_por_mes.csv")


def detectar_outliers(registros):
    """Avisa si algún cliente tiene una venta muy por encima de su mediana
    habitual — suele ser señal de un error de carga o de fecha/cliente."""
    por_cliente = {}
    for r in registros:
        por_cliente.setdefault(r["c"], []).append(r)

    avisos = []
    for cliente, ventas in por_cliente.items():
        montos = [v["t"] for v in ventas]
        if len(montos) < 3:
            continue
        med = median(montos)
        if med <= 0:
            continue
        for v in ventas:
            if v["t"] > med * 5:
                avisos.append((cliente, v["f"], v["t"], med))

    if avisos:
        print("\n⚠ Posibles ventas atípicas (más de 5x la mediana habitual del cliente):")
        for cliente, fecha, monto, med in avisos:
            print(f"   {cliente:30s} {fecha}  {pesos(monto):>17}   (habitual ~{pesos(med)})")
        print("   Revisá estas facturas en el sistema antes de sacar conclusiones del ranking.\n")
    else:
        print("\n✓ No se detectaron ventas atípicas evidentes.\n")


def generar_js(registros, salida):
    contenido = (
        "// Generado automáticamente por convertir_reparto.py\n"
        "// No editar manualmente — modificar el Excel exportado de Don Gestión y volver a correr el script\n"
        f"const RAW_REPARTO = {json.dumps(registros, ensure_ascii=False, indent=2)};\n"
    )
    Path(salida).write_text(contenido, encoding="utf-8")
    total = sum(r["t"] for r in registros)
    clientes = len(set(r["c"] for r in registros))
    print(f"✓ {len(registros)} ventas de reparto exportadas → {salida}")
    if registros:
        print(f"  Rango: {registros[0]['f']} a {registros[-1]['f']}")
    print(f"  Clientes únicos: {clientes}")
    print(f"  Total vendido: {pesos(total)}")


if __name__ == "__main__":
    archivo = encontrar_archivo()
    if not archivo:
        print("✗ No se encontró ningún Excel de reparto en esta carpeta.")
        print(f"  El archivo tiene que llamarse alguno de estos: {', '.join(CANDIDATOS_EXCEL)}")
    else:
        print(f"Leyendo '{archivo}'...")
        registros = leer_excel(archivo)
        totales_por_mes(registros)
        detectar_outliers(registros)
        generar_js(registros, ARCHIVO_SALIDA)
        print("\n  Ahora abrí dashboard.html en tu browser 🚀")