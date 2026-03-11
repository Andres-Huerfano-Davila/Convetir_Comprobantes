from __future__ import annotations

import streamlit as st

from processor import a_excel_bytes, construir_hojas, construir_resumen, consolidar_archivos

st.set_page_config(page_title="Gasto SAP PCP0", page_icon="📄", layout="wide")

st.title("📄 Generador de Gasto / Provisión desde SAP PCP0")
st.caption("Carga uno o varios TXT y genera las hojas Gasto, Provisión, Totales_Cuenta y Totales_CeCo.")

archivos = st.file_uploader(
    "Sube uno o varios TXT exportados desde SAP",
    type=["txt"],
    accept_multiple_files=True,
)

if not archivos:
    st.info("Carga al menos un archivo TXT para empezar.")
    st.stop()

df_base = consolidar_archivos(archivos)
if df_base.empty:
    st.error("No encontré registros válidos en los archivos cargados.")
    st.stop()

hojas = construir_hojas(df_base)
resumen = construir_resumen(df_base, hojas)

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

excel_bytes = a_excel_bytes(hojas, resumen=resumen)

st.download_button(
    "⬇️ Descargar Excel",
    data=excel_bytes,
    file_name="resultado_gasto_provision.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
