from __future__ import annotations

import traceback
from pathlib import Path

import streamlit as st

from processor import process_archivos

st.set_page_config(page_title="Gasto SAP PCP0", page_icon="📄", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
PDF_PATH = BASE_DIR / "Instructivo.pdf"

st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(180deg, #fff8f2 0%, #fff1e6 100%);
        }
        .block-container {
            padding-top: 1.4rem;
            padding-bottom: 1rem;
        }
        [data-testid="stSidebar"] {
            background: #fff3e6;
        }
        h1, h2, h3 {
            color: #9a4b00;
        }
        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.72);
            border: 1px solid #ffd2ad;
            border-radius: 14px;
            padding: 0.4rem 0.8rem;
        }
        div.stButton > button[kind="primary"] {
            background: #f28c28;
            border: 1px solid #f28c28;
            color: white;
        }
        div.stButton > button[kind="primary"]:hover {
            background: #dd7d1f;
            border-color: #dd7d1f;
            color: white;
        }
        div.stDownloadButton > button {
            border: 1px solid #f28c28;
            color: #9a4b00;
        }
        .marca-app {
            margin-top: 0.3rem;
            margin-bottom: 0.8rem;
            color: #b25a00;
            font-weight: 600;
        }
        .footer-jmc {
            text-align: center;
            margin-top: 1.2rem;
            padding: 0.7rem 0 0.2rem 0;
            color: #a14e00;
            font-weight: 600;
            border-top: 1px solid #ffd2ad;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📄 Generador de Gasto / Provisión desde SAP PCP0")
st.markdown('<div class="marca-app">Creado por Andrés Huérfano Dávila - Nómina JMC</div>', unsafe_allow_html=True)
st.caption(
    "Carga uno o varios archivos TXT, XLSX o XLS exportados desde SAP PCP0 y genera un Excel con Gasto, Provisión, Totales por Cuenta y Totales por CeCo."
)

if "resultado" not in st.session_state:
    st.session_state.resultado = None

with st.sidebar:
    st.subheader("Opciones")
    st.caption("Creado por Andrés Huérfano Dávila - Nómina JMC")
    preview_limit = st.number_input(
        "Filas de vista previa por hoja",
        min_value=10,
        max_value=5000,
        value=100,
        step=10,
        help="Solo afecta lo que ves en pantalla. El Excel se genera completo.",
    )
    st.info(
        "Para archivos gigantes, la app muestra solo una muestra en pantalla para no colgarse. "
        "El archivo descargable sí va completo."
    )
    if PDF_PATH.exists():
        st.download_button(
            "⬇️ Descargar instructivo PDF",
            data=PDF_PATH.read_bytes(),
            file_name="Instructivo.pdf",
            mime="application/pdf",
            width="stretch",
        )

archivos = st.file_uploader(
    "Sube uno o varios archivos TXT, XLSX o XLS exportados desde SAP",
    type=["txt", "xlsx", "xls"],
    accept_multiple_files=True,
)

col_a, col_b = st.columns([1, 1])
procesar = col_a.button("Procesar archivos", width="stretch", type="primary")
limpiar = col_b.button("Limpiar resultado", width="stretch")

if limpiar:
    st.session_state.resultado = None
    st.rerun()

if procesar:
    if not archivos:
        st.warning("Primero carga al menos un archivo TXT, XLSX o XLS.")
    else:
        barra = st.progress(0, text="Preparando...")
        estado = st.empty()
        try:
            def actualizar_progreso(ratio: float, mensaje: str) -> None:
                barra.progress(int(ratio * 100), text=mensaje)
                estado.info(mensaje)

            resultado = process_archivos(
                archivos,
                progress_callback=actualizar_progreso,
                preview_limit=int(preview_limit),
            )

            barra.progress(100, text="Proceso terminado")
            estado.success("Proceso finalizado correctamente.")
            st.session_state.resultado = resultado
        except Exception:
            st.session_state.resultado = None
            barra.empty()
            estado.empty()
            st.error("La app encontró un error al procesar los archivos.")
            st.code(traceback.format_exc())

resultado = st.session_state.resultado

manual_tab, uso_tab = st.tabs(["Instructivo", "Procesamiento"])

with manual_tab:
    st.subheader("Cómo preparar el comprobante desde SAP")
    st.markdown(
        """
        1. Filtra las ejecuciones por el estado de contabilización de nómina `PP`.
        2. Marca todos los comprobantes y entra a **Visualizar posiciones del documento**.
        3. Si quieres trabajar por subtotales, usa la columna **Texto expl.CC-nómina** y aplica **Subtotal**.
        4. Ajusta el ancho de las columnas antes de exportar.
        5. Descarga el archivo en **fichero local**, formato **TXT** y opción **No convertido**.
        6. Carga ese archivo en esta app y procesa.
        """
    )
    st.warning(
        "Si subes un TXT resumido, tabulado o exportado en otro formato, la app te avisará que la estructura del archivo no es la esperada."
    )
    if PDF_PATH.exists():
        st.download_button(
            "Descargar instructivo",
            data=PDF_PATH.read_bytes(),
            file_name="Instructivo.pdf",
            mime="application/pdf",
            width="content",
        )

with uso_tab:
    if not resultado:
        st.info("Carga tus archivos y pulsa 'Procesar archivos'.")
        st.markdown('<div class="footer-jmc">Creado por Andrés Huérfano Dávila - Nómina JMC</div>', unsafe_allow_html=True)
        st.stop()

    meta = resultado["meta"]
    resumen = resultado["resumen"]
    incidencias = resultado["incidencias"]
    hojas = resultado["preview_hojas"]
    excel_bytes = resultado["excel_bytes"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros válidos", f"{meta['registros_validos']:,}")
    c2.metric("Gasto", f"{meta['filas_gasto']:,}")
    c3.metric("Provisión", f"{meta['filas_provision']:,}")
    c4.metric("Archivos con incidencia", f"{meta['archivos_con_error']:,}")

    if meta["archivos_con_error"] > 0:
        st.warning(
            "Uno o más archivos no se procesaron completamente. Revisa la sección 'Incidencias' para ver si la estructura del archivo no es la esperada."
        )

    st.subheader("Resumen")
    st.dataframe(resumen, width="stretch", hide_index=True)

    if not incidencias.empty:
        st.subheader("Incidencias")
        st.dataframe(incidencias, width="stretch", hide_index=True)

    if meta["registros_validos"] == 0:
        st.error(
            "No se generaron registros válidos. Lo más probable es que la estructura del archivo no sea la esperada para este proceso."
        )
        st.download_button(
            "⬇️ Descargar Excel con resumen e incidencias",
            data=excel_bytes,
            file_name="resultado_gasto_provision.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.markdown('<div class="footer-jmc">Creado por Andrés Huérfano Dávila - Nómina JMC</div>', unsafe_allow_html=True)
        st.stop()

    if meta["gasto_truncado"] or meta["provision_truncado"]:
        st.warning(
            f"La vista previa está limitada a {meta['preview_limit']:,} filas por hoja para mantener la app estable. "
            "El Excel descargable sí contiene todo el detalle."
        )

    with st.expander("Hojas generadas"):
        st.write("**Gasto:**", ", ".join(meta["gasto_sheets"]))
        st.write("**Provisión:**", ", ".join(meta["provision_sheets"]))
        st.write("**Totales_Cuenta:**", ", ".join(meta["totales_cuenta_sheets"]))
        st.write("**Totales_CeCo:**", ", ".join(meta["totales_ceco_sheets"]))
        st.write("**Resumen:**", ", ".join(meta["resumen_sheets"]))
        st.write("**Incidencias:**", ", ".join(meta["incidencias_sheets"]))

    tab1, tab2, tab3, tab4 = st.tabs(["Gasto (vista previa)", "Provisión (vista previa)", "Totales_Cuenta", "Totales_CeCo"])

    with tab1:
        st.dataframe(hojas["Gasto"], width="stretch", hide_index=True)

    with tab2:
        st.dataframe(hojas["Provisión"], width="stretch", hide_index=True)

    with tab3:
        st.dataframe(hojas["Totales_Cuenta"], width="stretch", hide_index=True)

    with tab4:
        st.dataframe(hojas["Totales_CeCo"], width="stretch", hide_index=True)

    st.download_button(
        "⬇️ Descargar Excel completo",
        data=excel_bytes,
        file_name="resultado_gasto_provision.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

    with st.expander("Notas para archivos muy grandes"):
        st.markdown(
            """
            - La app procesa TXT y Excel en una sola pasada por archivo.
            - En pantalla solo se muestra una vista previa para no colgar el navegador.
            - Si una hoja supera el límite de Excel, se parte automáticamente en varias hojas (`Gasto_2`, `Provisión_2`, etc.).
            - En `Totales_CeCo` se toma el número completo extraído de `Imputación CO`, por ejemplo `CeCos 1011703999` → `1011703999`.
            - Si el TXT no fue exportado en formato correcto, la app lo reporta en `Incidencias`.
            """
        )

    st.markdown('<div class="footer-jmc">Creado por Andrés Huérfano Dávila - Nómina JMC</div>', unsafe_allow_html=True)
