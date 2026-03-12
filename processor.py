from __future__ import annotations

import io
import math
import re
from collections import defaultdict
from typing import Callable, Iterable, Iterator

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

try:
    import xlrd  # type: ignore
except Exception:  # pragma: no cover
    xlrd = None

ENCODINGS = ["utf-8", "cp1252", "latin-1"]
MAX_EXCEL_ROWS = 1_048_576
DATA_ROWS_PER_SHEET = MAX_EXCEL_ROWS - 1

COLUMNAS_BASE = [
    "Número cuenta c/txt.",
    "Imputación CO",
    "Texto expl.CC-nómina",
    "Nº pers.",
    "Impte.Debe",
    "Impte.Haber",
    "Mon.",
]

COLUMNAS_GASTO = COLUMNAS_BASE + ["cta", "ceco", "Observación"]
COLUMNAS_PROVISION = COLUMNAS_BASE
COLUMNAS_TOTALES_CUENTA = ["Cuenta", "Nombre cuenta", "Tipo", "Total Neto (Debe - Haber)", "Cantidad registros"]
COLUMNAS_TOTALES_CECO = ["CeCo", "Tipo", "Total Impte.Debe", "Cantidad registros"]
COLUMNAS_RESUMEN = ["Métrica", "Valor"]
COLUMNAS_INCIDENCIAS = ["Archivo", "Estado", "Detalle"]

MAPEO_CECO = {
    "60": "101",
    "62": "102",
    "63": "103",
}

HEADER_SYNONYMS = {
    "numero cuenta c/txt": "Número cuenta c/txt.",
    "numero cuenta ctxt": "Número cuenta c/txt.",
    "nmero cuenta c/txt": "Número cuenta c/txt.",
    "imputacion co": "Imputación CO",
    "texto explcc-nomina": "Texto expl.CC-nómina",
    "texto explccnomina": "Texto expl.CC-nómina",
    "texto expl ccnomina": "Texto expl.CC-nómina",
    "texto expl c c-nomina": "Texto expl.CC-nómina",
    "n pers": "Nº pers.",
    "no pers": "Nº pers.",
    "imptedebe": "Impte.Debe",
    "impte debe": "Impte.Debe",
    "imptehaber": "Impte.Haber",
    "impte haber": "Impte.Haber",
    "mon": "Mon.",
}

ProgressCallback = Callable[[float, str], None]


class FileStructureError(ValueError):
    pass


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


def normalize_text(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip()
    return text


def normalizar_header(value) -> str:
    text = normalize_text(value).lower()
    text = (
        text.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("º", "o")
        .replace("°", "o")
    )
    text = re.sub(r"[\.:]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def limpiar_numero(valor) -> float:
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if isinstance(valor, float) and (math.isnan(valor) or math.isinf(valor)):
            return 0.0
        return float(valor)
    texto = normalize_text(valor).replace(".", "").replace(",", ".")
    if not texto:
        return 0.0
    try:
        return float(texto)
    except ValueError:
        return 0.0


_DIGIT_SPLIT_RE = re.compile(r"^(\d+)\s*(.*)$")
_CECO_SHORT_RE = re.compile(r"(?i)ceco?s?\s*([0-9]{3})")
_CECO_FULL_RE = re.compile(r"(?i)ceco?s?\s*([0-9]+)")


def separar_cuenta(valor: str) -> tuple[str, str, bool]:
    raw = normalize_text(valor)
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
    texto = normalize_text(imputacion)
    if not texto:
        return ""
    match = _CECO_SHORT_RE.search(texto)
    return match.group(1) if match else ""


def extraer_ceco_completo(imputacion: str) -> str:
    texto = normalize_text(imputacion)
    if not texto:
        return ""
    match = _CECO_FULL_RE.search(texto)
    if match:
        return match.group(1)
    digits = re.sub(r"\D", "", texto)
    return digits


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
    return getattr(archivo, "name", f"archivo_{idx}")


def _guess_file_type(name: str) -> str:
    lowered = name.lower()
    if lowered.endswith(".txt"):
        return "txt"
    if lowered.endswith(".xlsx"):
        return "xlsx"
    if lowered.endswith(".xls"):
        return "xls"
    return "txt"


def _canonize_headers(row_values: list) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, value in enumerate(row_values):
        normalized = normalizar_header(value)
        canonical = HEADER_SYNONYMS.get(normalized)
        if canonical and canonical not in mapping:
            mapping[canonical] = idx
    return mapping


def _validate_header_map(header_map: dict[str, int], file_label: str) -> None:
    faltantes = [c for c in COLUMNAS_BASE if c not in header_map]
    if faltantes:
        raise FileStructureError(
            f"La estructura del archivo no es la esperada en {file_label}. "
            f"Faltan estas columnas: {', '.join(faltantes)}"
        )


def _detect_txt_format(texto: str, file_label: str) -> str:
    non_empty = [line for line in texto.splitlines()[:200] if line.strip()]
    if not non_empty:
        raise FileStructureError(f"El archivo {file_label} está vacío o no contiene datos legibles.")

    if any(line.startswith("|") and line.count("|") >= 8 for line in non_empty):
        return "pipe"

    for line in non_empty:
        if "\t" in line:
            row = [part for part in line.split("\t") if normalize_text(part)]
            header_map = _canonize_headers(row)
            if len(header_map) >= 5:
                raise FileStructureError(
                    f"La estructura del archivo no es la esperada en {file_label}. "
                    "Parece un TXT tabulado o un reporte resumido por concepto. "
                    "Exporta desde SAP en Visualizar posiciones del documento y guarda el TXT en opción 'No convertido'."
                )

    raise FileStructureError(
        f"La estructura del archivo no es la esperada en {file_label}. "
        "La app espera el TXT exportado desde SAP PCP0 en formato 'No convertido'."
    )


def iter_txt_records(data: bytes, progress_callback: ProgressCallback | None, start_ratio: float, end_ratio: float, file_label: str) -> Iterator[dict]:
    texto = decode_bytes(data)
    fmt = _detect_txt_format(texto, file_label)
    if fmt != "pipe":
        raise FileStructureError(
            f"La estructura del archivo no es la esperada en {file_label}."
        )

    total_lineas_archivo = max(1, texto.count("\n") + 1)
    rows_found = 0
    for num_linea, linea in enumerate(io.StringIO(texto), start=1):
        if num_linea % 20000 == 0 or num_linea == 1:
            archivo_avance = num_linea / total_lineas_archivo
            ratio = start_ratio + ((end_ratio - start_ratio) * archivo_avance)
            emit_progress(progress_callback, ratio, f"Procesando {file_label}: línea {num_linea:,} de {total_lineas_archivo:,}")
        if not es_linea_datos(linea):
            continue
        partes = linea.split("|")[1:-1]
        if len(partes) < 7:
            continue
        rows_found += 1
        yield {
            "Número cuenta c/txt.": partes[0],
            "Imputación CO": partes[1],
            "Texto expl.CC-nómina": partes[2],
            "Nº pers.": partes[3],
            "Impte.Debe": partes[4],
            "Impte.Haber": partes[5],
            "Mon.": partes[6],
        }

    if rows_found == 0:
        raise FileStructureError(
            f"La estructura del archivo no es la esperada en {file_label}. "
            "No se detectaron líneas de detalle válidas. Verifica que el TXT haya sido exportado en formato 'No convertido'."
        )


def iter_xlsx_records(data: bytes, progress_callback: ProgressCallback | None, start_ratio: float, end_ratio: float, file_label: str) -> Iterator[dict]:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    total_rows = max(1, ws.max_row or 1)
    header_map = None
    seen_rows = 0
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            header_map = _canonize_headers(list(row))
            _validate_header_map(header_map, file_label)
            continue
        if row_idx % 5000 == 0 or row_idx == 2:
            ratio = start_ratio + ((end_ratio - start_ratio) * (row_idx / total_rows))
            emit_progress(progress_callback, ratio, f"Procesando {file_label}: fila {row_idx:,} de {total_rows:,}")
        assert header_map is not None
        values = {col: row[header_map[col]] if header_map[col] < len(row) else None for col in COLUMNAS_BASE}
        seen_rows += 1
        yield values
    wb.close()
    if seen_rows == 0:
        raise FileStructureError(f"La estructura del archivo no es la esperada en {file_label}. El Excel no contiene filas de detalle.")


def iter_xls_records(data: bytes, progress_callback: ProgressCallback | None, start_ratio: float, end_ratio: float, file_label: str) -> Iterator[dict]:
    if xlrd is None:
        raise ImportError("Para leer archivos .xls debes incluir xlrd en requirements.txt")
    book = xlrd.open_workbook(file_contents=data)
    sheet = book.sheet_by_index(0)
    total_rows = max(1, sheet.nrows)
    header_map = _canonize_headers(sheet.row_values(0))
    _validate_header_map(header_map, file_label)
    if sheet.nrows <= 1:
        raise FileStructureError(f"La estructura del archivo no es la esperada en {file_label}. El Excel no contiene filas de detalle.")
    for row_idx in range(1, sheet.nrows):
        if row_idx % 5000 == 0 or row_idx == 1:
            ratio = start_ratio + ((end_ratio - start_ratio) * (row_idx / total_rows))
            emit_progress(progress_callback, ratio, f"Procesando {file_label}: fila {row_idx:,} de {total_rows:,}")
        row = sheet.row_values(row_idx)
        values = {col: row[header_map[col]] if header_map[col] < len(row) else None for col in COLUMNAS_BASE}
        yield values


def iter_records_for_file(file_name: str, data: bytes, progress_callback: ProgressCallback | None, start_ratio: float, end_ratio: float):
    file_type = _guess_file_type(file_name)
    if file_type == "txt":
        return iter_txt_records(data, progress_callback, start_ratio, end_ratio, file_name)
    if file_type == "xlsx":
        return iter_xlsx_records(data, progress_callback, start_ratio, end_ratio, file_name)
    if file_type == "xls":
        return iter_xls_records(data, progress_callback, start_ratio, end_ratio, file_name)
    raise ValueError(f"Tipo de archivo no soportado: {file_name}")


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
    incidencias: list[list[str]] = []

    total_validos = 0
    total_gasto = 0
    total_provision = 0
    suma_gasto = 0.0
    suma_provision = 0.0
    lineas_leidas = 0
    archivos_ok = 0
    archivos_con_error = 0

    emit_progress(progress_callback, 0.01, "Preparando procesamiento...")

    for idx, archivo in enumerate(archivos_list, start=1):
        nombre = _safe_name(archivo, idx)
        tramo_inicio = 0.02 + ((idx - 1) / total_archivos) * 0.86
        tramo_fin = 0.02 + (idx / total_archivos) * 0.86
        emit_progress(progress_callback, tramo_inicio, f"Leyendo {nombre}...")

        data = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read()
        file_rows = 0

        try:
            for raw in iter_records_for_file(nombre, data, progress_callback, tramo_inicio, tramo_fin):
                lineas_leidas += 1
                file_rows += 1

                cuenta_raw = normalize_text(raw.get("Número cuenta c/txt."))
                imputacion = normalize_text(raw.get("Imputación CO"))
                texto_nomina = normalize_text(raw.get("Texto expl.CC-nómina"))
                nro_personal = normalize_text(raw.get("Nº pers."))
                debe = limpiar_numero(raw.get("Impte.Debe"))
                haber = limpiar_numero(raw.get("Impte.Haber"))
                moneda = normalize_text(raw.get("Mon."))

                cuenta_num, nombre_cuenta, es_subtotal = separar_cuenta(cuenta_raw)
                if es_subtotal:
                    continue
                if debe == 0 and haber == 0:
                    continue

                cuenta_completa = f"{cuenta_num} {nombre_cuenta}".strip()
                ceco_validacion = extraer_ceco_3d(imputacion)
                ceco_total = extraer_ceco_completo(imputacion)
                tipo = "Gasto" if ceco_total else "Provisión"

                neto_movimiento = round(debe - haber, 2)

                if tipo == "Gasto":
                    debe_salida = neto_movimiento
                    haber_salida = haber if haber != 0 else None
                    cta = cuenta_num[:2] if cuenta_num else ""
                    observacion = construir_observacion(cta, ceco_validacion)
                    row = [
                        cuenta_completa,
                        imputacion,
                        texto_nomina,
                        nro_personal,
                        debe_salida,
                        haber_salida,
                        moneda,
                        cta,
                        ceco_validacion,
                        observacion,
                    ]
                    gasto_writer.append(row)
                    _append_preview(gasto_preview, row, preview_limit)
                    total_gasto += 1
                    suma_gasto += debe_salida
                    total_base = neto_movimiento
                    totales_ceco[(ceco_total, tipo)][0] += debe_salida
                    totales_ceco[(ceco_total, tipo)][1] += 1
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
                    total_base = neto_movimiento

                totales_cuenta[(cuenta_num, nombre_cuenta, tipo)][0] += total_base
                totales_cuenta[(cuenta_num, nombre_cuenta, tipo)][1] += 1
                total_validos += 1

            if file_rows == 0:
                incidencias.append([
                    nombre,
                    "Sin datos válidos",
                    "El archivo fue leído pero no generó registros útiles. Verifica subtotales, columnas y contenido.",
                ])
                archivos_con_error += 1
            else:
                archivos_ok += 1

        except Exception as exc:
            incidencias.append([nombre, "No procesado", str(exc)])
            archivos_con_error += 1
            continue

    emit_progress(progress_callback, 0.90, "Construyendo totales...")

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
        key=lambda x: (x[1], x[0]),
    )

    totales_cuenta_writer = SheetWriter(wb, "Totales_Cuenta", COLUMNAS_TOTALES_CUENTA)
    for row in totales_cuenta_rows:
        totales_cuenta_writer.append(row)

    totales_ceco_writer = SheetWriter(wb, "Totales_CeCo", COLUMNAS_TOTALES_CECO)
    for row in totales_ceco_rows:
        totales_ceco_writer.append(row)

    incidencias_writer = SheetWriter(wb, "Incidencias", COLUMNAS_INCIDENCIAS)
    for row in incidencias:
        incidencias_writer.append(row)

    resumen_rows = [
        ["Archivos recibidos", total_archivos],
        ["Archivos procesados", archivos_ok],
        ["Archivos con incidencia", archivos_con_error],
        ["Registros leídos", lineas_leidas],
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

    emit_progress(progress_callback, 0.96, "Generando archivo Excel...")
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    emit_progress(progress_callback, 1.0, "Proceso terminado")

    gasto_preview_df = pd.DataFrame(gasto_preview, columns=COLUMNAS_GASTO)
    provision_preview_df = pd.DataFrame(provision_preview, columns=COLUMNAS_PROVISION)
    totales_cuenta_df = pd.DataFrame(totales_cuenta_rows, columns=COLUMNAS_TOTALES_CUENTA)
    totales_ceco_df = pd.DataFrame(totales_ceco_rows, columns=COLUMNAS_TOTALES_CECO)
    resumen_df = pd.DataFrame(resumen_rows, columns=COLUMNAS_RESUMEN)
    incidencias_df = pd.DataFrame(incidencias, columns=COLUMNAS_INCIDENCIAS)

    return {
        "excel_bytes": excel_buffer.getvalue(),
        "resumen": resumen_df,
        "incidencias": incidencias_df,
        "preview_hojas": {
            "Gasto": gasto_preview_df,
            "Provisión": provision_preview_df,
            "Totales_Cuenta": totales_cuenta_df,
            "Totales_CeCo": totales_ceco_df,
        },
        "meta": {
            "archivos_recibidos": total_archivos,
            "archivos_procesados": archivos_ok,
            "archivos_con_error": archivos_con_error,
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
            "incidencias_sheets": incidencias_writer.sheet_names,
        },
    }
