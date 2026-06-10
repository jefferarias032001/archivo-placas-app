# -*- coding: utf-8 -*-
"""
convertir_viajes.py — Lee TODOS los Excel de la carpeta de viajes nacionales
y genera data_viajes.js con la inteligencia de flota por placa:

  - Viajes reales = conteo de "Envio" (manifiesto) DISTINTO por placa
  - Fecha del último viaje y días sin cargar
  - Estado de flota:
      FIDELIZADA     cargó en los últimos 30 días
      EVENTUAL       cargó entre 31 y 90 días
      POR RECUPERAR  más de 90 días sin cargar
  - Rotación = viajes promedio por mes (desde su primer viaje)
  - Historial agrupado por ORIGEN → DESTINO → CLIENTE con conteo de viajes

USO:
  1. Edita RUTA_CARPETA_VIAJES con tu carpeta de OneDrive.
  2. Ejecuta:  python convertir_viajes.py
     (o simplemente actualizar.bat, que corre todo junto)

Requiere:  pip install openpyxl
"""

import json
import os
import sys
import glob
from datetime import datetime, date

# ============================================================
# ⚙️ CONFIGURACIÓN — EDITA ESTA LÍNEA CON TU CARPETA REAL
# Es la CARPETA que contiene los Excel divididos por meses.
# Ejemplo:
#   C:/Users/jarias/OneDrive - TU_EMPRESA/VIAJES NACIONALES
# ============================================================
RUTA_CARPETA_VIAJES = "C:/Users/jarias/OneDrive - TRACTOCAR LOGISTICS SAS/POWER BI JEFFER/ARCHIVOS/NACIONAL"

SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_viajes.js")

# Umbrales de estado (en días)
DIAS_FIDELIZADA = 30
DIAS_EVENTUAL = 90


def limpio(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s in {"-", "N/D", "nan", "NaT"} else s


def quitar_prefijo(v):
    """TCL.LZO222 -> LZO222 ; TCL.2261640 -> 2261640"""
    s = limpio(v)
    return s.split(".", 1)[1] if s.startswith("TCL.") else s


def a_fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = limpio(v)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def main():
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("❌ Falta openpyxl. Instálala con:  pip install openpyxl")
        sys.exit(1)

    if not os.path.isdir(RUTA_CARPETA_VIAJES):
        print("❌ No encuentro la carpeta de viajes:")
        print("   " + RUTA_CARPETA_VIAJES)
        print("   Edita RUTA_CARPETA_VIAJES al inicio de convertir_viajes.py")
        sys.exit(1)

    archivos = sorted(
        f for f in glob.glob(os.path.join(RUTA_CARPETA_VIAJES, "*.xls*"))
        if not os.path.basename(f).startswith("~$")  # ignora temporales de Excel
    )
    if not archivos:
        print("❌ La carpeta no tiene archivos .xlsx:")
        print("   " + RUTA_CARPETA_VIAJES)
        sys.exit(1)

    print(f"📂 Carpeta: {RUTA_CARPETA_VIAJES}")
    print(f"   {len(archivos)} archivo(s) encontrados\n")

    # ---- acumuladores ----
    # flota[placa] = {"envios": set, "rutas": {(o,d,cliente): {"envios": set, "ult": date}},
    #                 "ult": date, "pri": date, "tipologia": {tip: n}}
    flota = {}
    envios_globales = set()
    fecha_min, fecha_max = None, None
    filas_totales = 0

    COLS = ["Envio", "Placa", "Ciudad Origen", "Ciudad Destino", "Tipologia", "Fecha Creacion", "Cliente"]

    for ruta_archivo in archivos:
        nombre = os.path.basename(ruta_archivo)
        try:
            wb = load_workbook(ruta_archivo, read_only=True, data_only=True)
        except Exception as e:
            print(f"   ⚠️ {nombre}: no se pudo abrir ({e}) — lo salto")
            continue

        # detectar la hoja que tenga Envio y Placa
        hoja, encabezados = None, None
        for nh in wb.sheetnames:
            primera = next(wb[nh].iter_rows(max_row=1, values_only=True), None)
            if primera and "Envio" in primera and "Placa" in primera:
                hoja, encabezados = nh, list(primera)
                break
        if not hoja:
            print(f"   ⚠️ {nombre}: ninguna hoja tiene columnas Envio y Placa — lo salto")
            continue

        idx = {n: i for i, n in enumerate(encabezados) if n}
        faltan = [c for c in COLS if c not in idx]
        if faltan:
            print(f"   ⚠️ {nombre}: faltan columnas {faltan} — lo salto")
            continue

        n_filas = 0
        for fila in wb[hoja].iter_rows(min_row=2, values_only=True):
            def col(nombre_col):
                i = idx[nombre_col]
                return fila[i] if i < len(fila) else None

            placa = quitar_prefijo(col("Placa"))
            envio = quitar_prefijo(col("Envio"))
            if not placa or not envio:
                continue
            n_filas += 1
            filas_totales += 1

            fecha = a_fecha(col("Fecha Creacion"))
            origen = limpio(col("Ciudad Origen")).upper()
            destino = limpio(col("Ciudad Destino")).upper()
            cliente = limpio(col("Cliente")).upper()
            tip = limpio(col("Tipologia")).upper()

            if fecha:
                fecha_min = fecha if not fecha_min or fecha < fecha_min else fecha_min
                fecha_max = fecha if not fecha_max or fecha > fecha_max else fecha_max

            f = flota.setdefault(placa, {"envios": set(), "rutas": {}, "ult": None, "pri": None, "tipologia": {}})
            f["envios"].add(envio)
            envios_globales.add(envio)
            if tip:
                f["tipologia"][tip] = f["tipologia"].get(tip, 0) + 1
            if fecha:
                f["ult"] = fecha if not f["ult"] or fecha > f["ult"] else f["ult"]
                f["pri"] = fecha if not f["pri"] or fecha < f["pri"] else f["pri"]

            clave = (origen, destino, cliente)
            r = f["rutas"].setdefault(clave, {"envios": set(), "ult": None})
            r["envios"].add(envio)
            if fecha:
                r["ult"] = fecha if not r["ult"] or fecha > r["ult"] else r["ult"]

        print(f"   ✅ {nombre}: {n_filas:,} filas")

    if not flota:
        print("❌ No se pudo leer ningún dato.")
        sys.exit(1)

    # ---- consolidar ----
    hoy = date.today()
    placas_out = []
    for placa, f in flota.items():
        viajes = len(f["envios"])
        ult, pri = f["ult"], f["pri"]
        dias = (hoy - ult).days if ult else 9999
        if dias <= DIAS_FIDELIZADA:
            estado = "FIDELIZADA"
        elif dias <= DIAS_EVENTUAL:
            estado = "EVENTUAL"
        else:
            estado = "POR RECUPERAR"
        meses = max(((hoy - pri).days / 30.44), 1.0) if pri else 1.0
        tip = max(f["tipologia"], key=f["tipologia"].get) if f["tipologia"] else ""

        rutas = []
        for (o, d, c), r in f["rutas"].items():
            rutas.append({
                "o": o, "d": d, "c": c,
                "n": len(r["envios"]),
                "f": r["ult"].strftime("%Y-%m-%d") if r["ult"] else "",
            })
        rutas.sort(key=lambda x: -x["n"])

        placas_out.append({
            "placa": placa,
            "tip": tip,
            "viajes": viajes,
            "ultima": ult.strftime("%Y-%m-%d") if ult else "",
            "primera": pri.strftime("%Y-%m-%d") if pri else "",
            "dias": dias,
            "estado": estado,
            "meses": round(meses, 2),
            "rot": round(viajes / meses, 2),
            "rutas": rutas,
        })

    placas_out.sort(key=lambda x: -x["viajes"])

    paquete = {
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "desde": fecha_min.strftime("%Y-%m-%d") if fecha_min else "",
        "hasta": fecha_max.strftime("%Y-%m-%d") if fecha_max else "",
        "archivos": len(archivos),
        "filas": filas_totales,
        "totalViajes": len(envios_globales),
        "umbralFidelizada": DIAS_FIDELIZADA,
        "umbralEventual": DIAS_EVENTUAL,
        "placas": placas_out,
    }

    with open(SALIDA, "w", encoding="utf-8") as f:
        f.write("// Generado por convertir_viajes.py — NO editar a mano.\n")
        f.write("window.DATOS_VIAJES = ")
        json.dump(paquete, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")

    peso = os.path.getsize(SALIDA) / 1024 / 1024
    fid = sum(1 for p in placas_out if p["estado"] == "FIDELIZADA")
    eve = sum(1 for p in placas_out if p["estado"] == "EVENTUAL")
    rec = sum(1 for p in placas_out if p["estado"] == "POR RECUPERAR")
    print(f"\n✅ data_viajes.js generado — {peso:.1f} MB")
    print(f"   {len(placas_out):,} placas · {len(envios_globales):,} viajes · {fecha_min} → {fecha_max}")
    print(f"   Fidelizadas: {fid:,} · Eventuales: {eve:,} · Por recuperar: {rec:,}")


if __name__ == "__main__":
    main()
