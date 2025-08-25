import streamlit as st
from reporte_oa2 import procesar_oa2

st.title("Reporte OA-2 de Expedientes Judiciales")

st.markdown("Cargue los archivos Excel pasado y actual:")

file_pasado = st.file_uploader("Archivo PASADO", type=["xls", "xlsx"], key="pasado")
file_actual = st.file_uploader("Archivo ACTUAL", type=["xls", "xlsx"], key="actual")

if st.button("Procesar archivos"):
    if file_pasado and file_actual:
        output, comparaciones = procesar_oa2(file_pasado, file_actual)
        if output:
            st.success("Procesamiento completado.")
            st.download_button("Descargar reporte OA-2", data=output, file_name="Reporte_OA2.xlsx")
            st.markdown("### Resumen comparaciones:")
            for nombre, df in comparaciones.items():
                st.markdown(f"#### {nombre}")
                st.dataframe(df)
        else:
            st.error("Ocurrió un error al procesar los archivos.")
    else:
        st.warning("Por favor, cargue ambos archivos Excel.")
