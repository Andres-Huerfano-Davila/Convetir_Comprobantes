from __future__ import annotations

import traceback

import streamlit as st

from processor import a_excel_bytes, construir_hojas, construir_resumen, consolidar_archivos

st.set_page_config(page_title="Gasto SAP PCP0", page_icon="📄", layout="wide")

st.title("📄 Generador de Gasto / Provisión desde SAP PCP0")
st.caption("Carga uno o varios TXT y genera las hojas Gasto, Provisión, Totales_Cuenta y Totales_CeCo.")

if "resultado" not in st.session_state:
    st.session_state.resultado = None

archivos = st.file_uploader(
    "Sube uno o varios TXT exportados desde SAP",
    type=["txt"],
    accept_multiple_files=True,
)

col_a, col_b = st.columns([1, 3])
procesar = col_a.button("Procesar archivos", width="stretch", type="primary")
limpiar = col_b.button("Limpiar resultado", width="content")

if limpiar:
    st.session_state.resultado = None
    st.rerun()

if procesar:
    if not archivos:
        st.warning("Primero carga al menos un archivo TXT.")
    else:
        barra = st.progress(0, text="Iniciando proceso...")
        estado = st.empty()
        try:
            def actualizar_progreso(actual: int, total: int, nombre: str) -> None:
                if total <= 0:
                    barra.progress(10, text="Leyendo archivos...")
                    estado.info("Leyendo archivos...")
                    return
                progreso = int(10 + (actual / total) * 60)
                barra.progress(min(progreso, 70), text=f"Procesando archivo {actual}/{total}: {nombre}")
                estado.info(f"Procesando archivo {actual}/{total}: {nombre}")

            barra.progress(5, text="Leyendo archivos cargados...")
            df_base = consolidar_archivos(archivos, progress_callback=actualizar_progreso)

            barra.progress(75, text="Construyendo hojas...")
            estado.info("Construyendo hojas Gasto y Provisión...")
            hojas = construir_hojas(df_base)

            barra.progress(85, text="Construyendo resumen...")
            estado.info("Calculando resumen y totales...")
            resumen = construir_resumen(df_base, hojas)

            barra.progress(95, text="Generando Excel...")
            estado.info("Generando archivo Excel...")
            excel_bytes = a_excel_bytes(hojas, resumen=resumen)

            barra.progress(100, text="Proceso terminado")
            estado.success("Proceso finalizado correctamente.")

            st.session_state.resultado = {
                "df_base": df_base,
                "hojas": hojas,
                "resumen": resumen,
                "excel_bytes": excel_bytes,
            }
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


df_base = resultado["df_base"]
hojas = resultado["hojas"]
resumen = resultado["resumen"]
excel_bytes = resultado["excel_bytes"]

if df_base.empty:
    st.warning("No encontré registros válidos en los archivos cargados.")
    st.stop()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Registros válidos", f"{len(df_base):,}")
c2.metric("Gasto", f"{len(hojas['Gasto']):,}")
c3.metric("Provisión", f"{len(hojas['Provisión']):,}")
c4.metric("CeCos consolidados", f"{len(hojas['Totales_CeCo']):,}")

st.subheader("Resumen")
st.dataframe(resumen, width="stretch", hide_index=True)

tab1, tab2, tab3, tab4 = st.tabs(["Gasto", "Provisión", "Totales_Cuenta", "Totales_CeCo"])

with tab1:
    st.dataframe(hojas["Gasto"], width="stretch", hide_index=True)

with tab2:
    st.dataframe(hojas["Provisión"], width="stretch", hide_index=True)

with tab3:
    st.dataframe(hojas["Totales_Cuenta"], width="stretch", hide_index=True)

with tab4:
    st.dataframe(hojas["Totales_CeCo"], width="stretch", hide_index=True)

st.download_button(
    "⬇️ Descargar Excel",
    data=excel_bytes,
    file_name="resultado_gasto_provision.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)

with st.expander("Lógica aplicada"):
    st.markdown(
        """
        - Omite filas subtotal marcadas con `*`.
        - Omite filas donde Debe y Haber vienen vacíos.
        - En **Gasto**, convierte el Haber al Debe negativo con `Debe - Haber`.
        - En **Provisión**, deja Debe y Haber tal como vienen del TXT.
        - En **Gasto**, crea `cta`, `ceco` y `Observación`.
        - En **Provisión**, no genera `cta`, `ceco` ni `Observación`.
        - **Totales_Cuenta** incluye cuenta, nombre de cuenta y tipo.
        - **Totales_CeCo** consolida solo los registros de Gasto por CeCo.
        """
    )
