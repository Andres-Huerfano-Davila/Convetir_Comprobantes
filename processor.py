from __future__ import annotations

import io
import re
from collections import defaultdict
from typing import Callable, Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

ENCODINGS = ["utf-8", "cp1252", "latin-1"]
MAX_EXCEL_ROWS = 1_048_576
DATA_ROWS_PER_SHEET = MAX_EXCEL_ROWS - 1

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

COLUMNAS_TOTALES_CUENTA = ["Cuenta", "Nombre cuenta", "Tipo", "Total Impte.Debe", "Cantidad registros"]
COLUMNAS_TOTALES_CECO = ["CeCo", "Tipo", "Total Impte.Debe", "Cantidad registros"]
COLUMNAS_RESUMEN = ["Métrica", "Valor"]

MAPEO_CECO = {
    "60": "101",
    "62": "102",
    "63": "103",
}

ProgressCallback = Callable[[float, str], None]


class SheetWriter:
    def __init__(self, wb: Workbook, base_name: str, headers: list[str]):
        self.wb = wb
        self.base_name = base_name
        self.headers = headers
        self.sheet_index = 0
        self.data_rows = 0
        self.total_rows = 0
        self.sheet_names: list[str] = []
        self.ws = None
        self._new_sheet()

    def _header_cells(self):
        fill = PatternFill("solid", fgColor="1F4E78")
        font = Font(color="FFFFFF", bold=True)
        row = []
        for value in self.headers:
            cell = WriteOnlyCell(self.ws, value=value)
            cell.fill = fill
            cell.font = font
            row.append(cell)
        return row

    def _new_sheet(self):
        self.sheet_index += 1
        name = self.base_name if self.sheet_index == 1 else f"{self.base_name}_{self.sheet_index}"
        self.ws = self.wb.create_sheet(name[:31])
        self.sheet_names.append(self.ws.title)
        self.data_rows = 0
        self.ws.append(self._header_cells())

    def append(self, row: list):
        if self.data_rows >= DATA_ROWS_PER_SHEET:
            self._new_sheet()
        self.ws.append(row)
        self.data_rows += 1
        self.total_rows += 1


def emit_progress(callback: ProgressCallback | None, ratio: float, message: str) -> None:
    if callback:
        callback(max(0.0, min(1.0, ratio)), message)


def decode_bytes(data: bytes) -> str:
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


_DIGIT_SPLIT_RE = re.compile(r"^(\d+)\s*(.*)$")
_CECO_RE = re.compile(r"(?i)ceco?s?\s*([0-9]{3})")


def separar_cuenta(valor: str) -> tuple[str, str, bool]:
    raw = (valor or "").rstrip()
    es_subtotal = raw.lstrip().startswith("*")
    limpio = raw.lstrip(" *")
    match = _DIGIT_SPLIT_RE.match(limpio)
    if not match:
        return "", limpio.strip(), es_subtotal
    cuenta = match.group(1).strip()
    nombre = match.group(2).strip()
    return cuenta, nombre, es_subtotal


def es_linea_datos(linea: str) -> bool:
    if not linea.startswith("|"):
        return False
    core = linea.replace("|", "").strip()
    if not core:
        return False
    if set(core) <= {"-"}:
        return False
    if "Número cuenta c/txt." in linea or "Nmero cuenta c/txt." in linea:
        return False
    return linea.count("|") >= 8


def extraer_ceco_3d(imputacion: str) -> str:
    texto = (imputacion or "").strip()
    if not texto:
        return ""
    match = _CECO_RE.search(texto)
    return match.group(1) if match else ""


def construir_observacion(cta: str, ceco: str) -> str:
    esperado = MAPEO_CECO.get(cta)
    if esperado and esperado == ceco:
        return "CeCo coincide con cuenta"
    if esperado:
        return "El ceco No coincide con la cuenta de gasto"
    return "Regla no definida para la cuenta"


def _append_preview(target: list[list], row: list, preview_limit: int) -> None:
    if len(target) < preview_limit:
        target.append(row)


def _safe_name(archivo, idx: int) -> str:
    return getattr(archivo, "name", f"archivo_{idx}.txt")


def process_archivos(
    archivos: Iterable,
    progress_callback: ProgressCallback | None = None,
    preview_limit: int = 1000,
) -> dict:
    archivos_list = list(archivos)
    total_archivos = len(archivos_list)
    if total_archivos == 0:
        raise ValueError("No se recibieron archivos para procesar.")

    wb = Workbook(write_only=True)

    gasto_writer = SheetWriter(wb, "Gasto", COLUMNAS_GASTO)
    provision_writer = SheetWriter(wb, "Provisión", COLUMNAS_PROVISION)

    gasto_preview: list[list] = []
    provision_preview: list[list] = []

    totales_cuenta: dict[tuple[str, str, str], list[float | int]] = defaultdict(lambda: [0.0, 0])
    totales_ceco: dict[tuple[str, str], list[float | int]] = defaultdict(lambda: [0.0, 0])

    total_validos = 0
    total_gasto = 0
    total_provision = 0
    suma_gasto = 0.0
    suma_provision = 0.0
    lineas_leidas = 0

    emit_progress(progress_callback, 0.01, "Preparando procesamiento...")

    for idx, archivo in enumerate(archivos_list, start=1):
        nombre = _safe_name(archivo, idx)
        emit_progress(progress_callback, 0.02 + ((idx - 1) / total_archivos) * 0.90, f"Leyendo {nombre}...")

        data = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read()
        texto = decode_bytes(data)
        total_lineas_archivo = max(1, texto.count("\n") + 1)

        for num_linea, linea in enumerate(io.StringIO(texto), start=1):
            lineas_leidas += 1
            if num_linea % 20000 == 0 or num_linea == 1:
                archivo_base = (idx - 1) / total_archivos
                archivo_avance = num_linea / total_lineas_archivo
                ratio = 0.02 + ((archivo_base + archivo_avance / total_archivos) * 0.90)
                emit_progress(progress_callback, ratio, f"Procesando {nombre}: línea {num_linea:,} de {total_lineas_archivo:,}")

            if not es_linea_datos(linea):
                continue

            partes = linea.split("|")[1:-1]
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

            cuenta_completa = f"{cuenta_num} {nombre_cuenta}".strip()
            ceco_extraido = extraer_ceco_3d(imputacion)
            tipo = "Gasto" if ceco_extraido else "Provisión"

            if tipo == "Gasto":
                debe_salida = round(debe - haber, 2)
                haber_salida = haber if haber != 0 else None
                cta = cuenta_num[:2] if cuenta_num else ""
                observacion = construir_observacion(cta, ceco_extraido)
                row = [
                    cuenta_completa,
                    imputacion,
                    texto_nomina,
                    nro_personal,
                    debe_salida,
                    haber_salida,
                    moneda,
                    cta,
                    ceco_extraido,
                    observacion,
                ]
                gasto_writer.append(row)
                _append_preview(gasto_preview, row, preview_limit)
                total_gasto += 1
                suma_gasto += debe_salida
                total_base = debe_salida
                totales_ceco[(ceco_extraido, tipo)][0] += debe_salida
                totales_ceco[(ceco_extraido, tipo)][1] += 1
            else:
                debe_salida = debe if debe != 0 else None
                haber_salida = haber if haber != 0 else None
                row = [
                    cuenta_completa,
                    imputacion,
                    texto_nomina,
                    nro_personal,
                    debe_salida,
                    haber_salida,
                    moneda,
                ]
                provision_writer.append(row)
                _append_preview(provision_preview, row, preview_limit)
                total_provision += 1
                suma_provision += debe if debe != 0 else 0.0
                total_base = debe if debe != 0 else 0.0

            totales_cuenta[(cuenta_num, nombre_cuenta, tipo)][0] += total_base
            totales_cuenta[(cuenta_num, nombre_cuenta, tipo)][1] += 1
            total_validos += 1

    emit_progress(progress_callback, 0.93, "Construyendo totales...")

    totales_cuenta_rows = sorted(
        [
            [cuenta, nombre, tipo, round(total, 2), cantidad]
            for (cuenta, nombre, tipo), (total, cantidad) in totales_cuenta.items()
        ],
        key=lambda x: (x[2], x[0], x[1]),
    )
    totales_ceco_rows = sorted(
        [
            [ceco, tipo, round(total, 2), cantidad]
            for (ceco, tipo), (total, cantidad) in totales_ceco.items()
        ],
        key=lambda x: (x[0], x[1]),
    )

    totales_cuenta_writer = SheetWriter(wb, "Totales_Cuenta", COLUMNAS_TOTALES_CUENTA)
    for row in totales_cuenta_rows:
        totales_cuenta_writer.append(row)

    totales_ceco_writer = SheetWriter(wb, "Totales_CeCo", COLUMNAS_TOTALES_CECO)
    for row in totales_ceco_rows:
        totales_ceco_writer.append(row)

    resumen_rows = [
        ["Archivos procesados", total_archivos],
        ["Líneas leídas", lineas_leidas],
        ["Registros válidos", total_validos],
        ["Filas Gasto", total_gasto],
        ["Filas Provisión", total_provision],
        ["Total Impte.Debe Gasto", round(suma_gasto, 2)],
        ["Total Impte.Debe Provisión", round(suma_provision, 2)],
        ["Hojas Gasto generadas", len(gasto_writer.sheet_names)],
        ["Hojas Provisión generadas", len(provision_writer.sheet_names)],
    ]

    resumen_writer = SheetWriter(wb, "Resumen", COLUMNAS_RESUMEN)
    for row in resumen_rows:
        resumen_writer.append(row)

    emit_progress(progress_callback, 0.97, "Generando archivo Excel...")
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    emit_progress(progress_callback, 1.0, "Proceso terminado")

    gasto_preview_df = pd.DataFrame(gasto_preview, columns=COLUMNAS_GASTO)
    provision_preview_df = pd.DataFrame(provision_preview, columns=COLUMNAS_PROVISION)
    totales_cuenta_df = pd.DataFrame(totales_cuenta_rows, columns=COLUMNAS_TOTALES_CUENTA)
    totales_ceco_df = pd.DataFrame(totales_ceco_rows, columns=COLUMNAS_TOTALES_CECO)
    resumen_df = pd.DataFrame(resumen_rows, columns=COLUMNAS_RESUMEN)

    return {
        "excel_bytes": excel_buffer.getvalue(),
        "resumen": resumen_df,
        "preview_hojas": {
            "Gasto": gasto_preview_df,
            "Provisión": provision_preview_df,
            "Totales_Cuenta": totales_cuenta_df,
            "Totales_CeCo": totales_ceco_df,
        },
        "meta": {
            "archivos_procesados": total_archivos,
            "lineas_leidas": lineas_leidas,
            "registros_validos": total_validos,
            "filas_gasto": total_gasto,
            "filas_provision": total_provision,
            "total_gasto": round(suma_gasto, 2),
            "total_provision": round(suma_provision, 2),
            "preview_limit": preview_limit,
            "gasto_truncado": total_gasto > preview_limit,
            "provision_truncado": total_provision > preview_limit,
            "gasto_sheets": gasto_writer.sheet_names,
            "provision_sheets": provision_writer.sheet_names,
            "totales_cuenta_sheets": totales_cuenta_writer.sheet_names,
            "totales_ceco_sheets": totales_ceco_writer.sheet_names,
            "resumen_sheets": resumen_writer.sheet_names,
        },
    }
