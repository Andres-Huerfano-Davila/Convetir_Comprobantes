from __future__ import annotations

import io
import re
from typing import Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ENCODINGS = ["utf-8", "cp1252", "latin-1"]

COLUMNAS_GASTO = [
    "Número cuenta c/txt.",
    "Imputación CO",
    "Texto expl.CC-nómina",
    "Nº pers.",
    "Impte.Debe",
    "Impte.Haber",
    "Mon.",
    "cta",
    "ceco",
    "Observación",
]

COLUMNAS_PROVISION = [
    "Número cuenta c/txt.",
    "Imputación CO",
    "Texto expl.CC-nómina",
    "Nº pers.",
    "Impte.Debe",
    "Impte.Haber",
    "Mon.",
]

MAPEO_CECO = {
    "60": "101",
    "62": "102",
    "63": "103",
}


def leer_texto_bytes(data: bytes) -> str:
    for enc in ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="ignore")


def limpiar_numero(valor: str) -> float:
    texto = (valor or "").strip().replace(".", "").replace(",", ".")
    if not texto:
        return 0.0
    try:
        return float(texto)
    except ValueError:
        return 0.0


def separar_cuenta(valor: str) -> tuple[str, str, bool]:
    raw = (valor or "").rstrip()
    es_subtotal = raw.lstrip().startswith("*")
    limpio = raw.lstrip(" *")

    match = re.match(r"^(\d+)\s+(.*)$", limpio)
    if match:
        return match.group(1).strip(), match.group(2).strip(), es_subtotal

    match = re.match(r"^(\d+)$", limpio)
    if match:
        return match.group(1).strip(), "", es_subtotal

    return "", limpio.strip(), es_subtotal


def es_linea_datos(linea: str) -> bool:
    if not linea.startswith("|"):
        return False
    if set(linea.replace("|", "").strip()) <= {"-"}:
        return False
    if "Número cuenta c/txt." in linea or "Nmero cuenta c/txt." in linea:
        return False
    return len(linea.split("|")) >= 9


def extraer_ceco_3d(imputacion: str) -> str:
    texto = (imputacion or "").strip()
    if not texto:
        return ""
    match = re.search(r"(?i)ceco?s?\s*([0-9]{3})", texto)
    if match:
        return match.group(1)
    return ""


def construir_observacion(cta: str, ceco: str) -> str:
    esperado = MAPEO_CECO.get(cta)
    if esperado and esperado == ceco:
        return "CeCo coincide con cuenta"
    if esperado:
        return "El ceco No coincide con la cuenta de gasto"
    return "Regla no definida para la cuenta"


def parsear_txt_contenido(texto: str, nombre_archivo: str) -> pd.DataFrame:
    registros: list[dict] = []

    for linea in texto.splitlines():
        if not es_linea_datos(linea):
            continue

        partes = [p for p in linea.split("|")[1:-1]]
        if len(partes) < 7:
            continue

        cuenta_raw = partes[0]
        imputacion = partes[1].strip()
        texto_nomina = partes[2].strip()
        nro_personal = partes[3].strip()
        debe_txt = partes[4].strip()
        haber_txt = partes[5].strip()
        moneda = partes[6].strip()

        cuenta_num, nombre_cuenta, es_subtotal = separar_cuenta(cuenta_raw)
        if es_subtotal:
            continue

        debe = limpiar_numero(debe_txt)
        haber = limpiar_numero(haber_txt)

        if debe == 0 and haber == 0:
            continue

        ceco_extraido = extraer_ceco_3d(imputacion)
        tipo = "Gasto" if ceco_extraido else "Provisión"

        if tipo == "Gasto":
            debe_salida = round(debe - haber, 2)
            haber_salida = haber if haber != 0 else None
            cta = cuenta_num[:2] if cuenta_num else ""
            ceco = ceco_extraido
            observacion = construir_observacion(cta, ceco)
            total_base = debe_salida
        else:
            debe_salida = debe if debe != 0 else None
            haber_salida = haber if haber != 0 else None
            cta = ""
            ceco = ""
            observacion = ""
            total_base = debe

        registros.append(
            {
                "archivo": nombre_archivo,
                "Número cuenta c/txt.": f"{cuenta_num} {nombre_cuenta}".strip(),
                "Imputación CO": imputacion,
                "Texto expl.CC-nómina": texto_nomina,
                "Nº pers.": nro_personal,
                "Impte.Debe": debe_salida,
                "Impte.Haber": haber_salida,
                "Mon.": moneda,
                "cta": cta,
                "ceco": ceco,
                "Observación": observacion,
                "_cuenta_num": cuenta_num,
                "_nombre_cuenta": nombre_cuenta,
                "_tipo": tipo,
                "_total_base": total_base,
                "_ceco": ceco_extraido,
            }
        )

    return pd.DataFrame(registros)


def consolidar_archivos(archivos: Iterable) -> pd.DataFrame:
    marcos = []
    for archivo in archivos:
        data = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read()
        texto = leer_texto_bytes(data)
        nombre = getattr(archivo, "name", "archivo.txt")
        df = parsear_txt_contenido(texto, nombre)
        if not df.empty:
            marcos.append(df)
    if not marcos:
        return pd.DataFrame()
    return pd.concat(marcos, ignore_index=True)


def construir_hojas(df_base: pd.DataFrame) -> dict[str, pd.DataFrame]:
    if df_base.empty:
        return {
            "Gasto": pd.DataFrame(columns=COLUMNAS_GASTO),
            "Provisión": pd.DataFrame(columns=COLUMNAS_PROVISION),
            "Totales_Cuenta": pd.DataFrame(columns=["Cuenta", "Nombre cuenta", "Tipo", "Total Impte.Debe", "Cantidad registros"]),
            "Totales_CeCo": pd.DataFrame(columns=["CeCo", "Tipo", "Total Impte.Debe", "Cantidad registros"]),
        }

    gasto = (
        df_base[df_base["_tipo"] == "Gasto"][COLUMNAS_GASTO]
        .sort_values(["Número cuenta c/txt.", "Imputación CO", "Texto expl.CC-nómina", "Nº pers."])
        .reset_index(drop=True)
    )

    provision = (
        df_base[df_base["_tipo"] == "Provisión"][COLUMNAS_PROVISION]
        .sort_values(["Número cuenta c/txt.", "Texto expl.CC-nómina", "Nº pers."])
        .reset_index(drop=True)
    )

    totales_cuenta = (
        df_base.groupby(["_cuenta_num", "_nombre_cuenta", "_tipo"], dropna=False, as_index=False)
        .agg(**{"Total Impte.Debe": ("_total_base", "sum"), "Cantidad registros": ("_total_base", "size")})
        .rename(columns={"_cuenta_num": "Cuenta", "_nombre_cuenta": "Nombre cuenta", "_tipo": "Tipo"})
        .sort_values(["Tipo", "Cuenta", "Nombre cuenta"])
        .reset_index(drop=True)
    )

    totales_ceco = (
        df_base[df_base["_tipo"] == "Gasto"]
        .groupby(["_ceco", "_tipo"], dropna=False, as_index=False)
        .agg(**{"Total Impte.Debe": ("_total_base", "sum"), "Cantidad registros": ("_total_base", "size")})
        .rename(columns={"_ceco": "CeCo", "_tipo": "Tipo"})
        .sort_values(["CeCo", "Tipo"])
        .reset_index(drop=True)
    )

    return {
        "Gasto": gasto,
        "Provisión": provision,
        "Totales_Cuenta": totales_cuenta,
        "Totales_CeCo": totales_ceco,
    }


def construir_resumen(df_base: pd.DataFrame, hojas: dict[str, pd.DataFrame]) -> pd.DataFrame:
    if df_base.empty:
        return pd.DataFrame(columns=["Métrica", "Valor"])

    total_gasto = float(hojas["Gasto"]["Impte.Debe"].fillna(0).sum()) if not hojas["Gasto"].empty else 0.0
    total_provision = float(hojas["Provisión"]["Impte.Debe"].fillna(0).sum()) if not hojas["Provisión"].empty else 0.0
    total_ceco = float(hojas["Totales_CeCo"]["Total Impte.Debe"].fillna(0).sum()) if not hojas["Totales_CeCo"].empty else 0.0

    return pd.DataFrame(
        {
            "Métrica": [
                "Archivos procesados",
                "Registros válidos",
                "Filas Gasto",
                "Filas Provisión",
                "Total Impte.Debe Gasto",
                "Total Impte.Debe Provisión",
                "Total consolidado por CeCo",
            ],
            "Valor": [
                int(df_base["archivo"].nunique()),
                int(len(df_base)),
                int(len(hojas["Gasto"])),
                int(len(hojas["Provisión"])),
                round(total_gasto, 2),
                round(total_provision, 2),
                round(total_ceco, 2),
            ],
        }
    )


def _aplicar_estilo_hoja(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E1F2")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_gray)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    encabezados = [cell.value for cell in ws[1]]
    cols_numericas = {"Impte.Debe", "Impte.Haber", "Total Impte.Debe", "Cantidad registros", "Valor"}

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center")
            if encabezados[cell.column - 1] in cols_numericas and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0;[Red]-#,##0'

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def a_excel_bytes(hojas: dict[str, pd.DataFrame], resumen: pd.DataFrame | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    wb.remove(ws)

    if resumen is not None:
        ws_resumen = wb.create_sheet("Resumen")
        ws_resumen.append(list(resumen.columns))
        for row in resumen.itertuples(index=False):
            ws_resumen.append(list(row))
        _aplicar_estilo_hoja(ws_resumen)

    orden_hojas = ["Gasto", "Provisión", "Totales_Cuenta", "Totales_CeCo"]
    for nombre_hoja in orden_hojas:
        df = hojas.get(nombre_hoja, pd.DataFrame())
        ws = wb.create_sheet(nombre_hoja[:31])
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append(list(row))
        _aplicar_estilo_hoja(ws)

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()
