from __future__ import annotations

import traceback

import streamlit as st

from processor import process_archivos

st.set_page_config(page_title="Gasto SAP PCP0", page_icon="📄", layout="wide")

st.title("📄 Generador de Gasto / Provisión desde SAP PCP0")
st.caption(
    "Carga uno o varios TXT exportados desde SAP PCP0 y genera un Excel con Gasto, Provisión, Totales por Cuenta y Totales por CeCo."
)

if "resultado" not in st.session_state:
    st.session_state.resultado = None

with st.sidebar:
    st.subheader("Opciones")
    preview_limit = st.number_input(
        "Filas de vista previa por hoja",
        min_value=100,
        max_value=5000,
        value=1000,
        step=100,
        help="Solo afecta lo que ves en pantalla. El Excel se genera completo.",
    )
    st.info(
        "Para archivos gigantes, la app muestra solo una muestra en pantalla para no colgarse. "
        "El archivo descargable sí va completo."
    )

archivos = st.file_uploader(
    "Sube uno o varios TXT exportados desde SAP",
    type=["txt"],
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
        st.warning("Primero carga al menos un archivo TXT.")
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

if not resultado:
    st.info("Carga tus TXT y pulsa 'Procesar archivos'.")
    st.stop()

meta = resultado["meta"]
resumen = resultado["resumen"]
hojas = resultado["preview_hojas"]
excel_bytes = resultado["excel_bytes"]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Registros válidos", f"{meta['registros_validos']:,}")
c2.metric("Gasto", f"{meta['filas_gasto']:,}")
c3.metric("Provisión", f"{meta['filas_provision']:,}")
c4.metric("Líneas leídas", f"{meta['lineas_leidas']:,}")

st.subheader("Resumen")
st.dataframe(resumen, width="stretch", hide_index=True)

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
        - La app procesa los TXT grandes en una sola pasada.
        - En pantalla solo se muestra una vista previa para no colgar el navegador.
        - Si una hoja supera el límite de Excel, se parte automáticamente en varias hojas (`Gasto_2`, `Provisión_2`, etc.).
        - Para cargas enormes, la generación del Excel puede tomar varios minutos.
        """
    )
