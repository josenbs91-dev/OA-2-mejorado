import streamlit as st
from reporte_oa2 import procesar_oa2

st.set_page_config(page_title="Análisis OA-2", layout="wide")

st.title("Análisis Comparativo de Expedientes Judiciales (OA-2)")

# Carga de archivos
file_pasado = st.file_uploader("Cargar archivo PASADO", type=["xlsx"])
file_actual = st.file_uploader("Cargar archivo ACTUAL", type=["xlsx"])

if file_pasado and file_actual:
    st.success("Archivos cargados correctamente.")
    
    if st.button("Generar Reporte OA-2"):
        output, comparacion = procesar_oa2(file_pasado, file_actual)
        
        if output:
            st.success("Reporte generado correctamente.")
            
            # Descarga
            st.download_button(
                label="Descargar Excel OA-2",
                data=output,
                file_name="Reporte_OA2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Mostrar tabla comparativa
            st.subheader("Comparación por Cuenta")
            st.dataframe(comparacion)
        else:
            st.error("Ocurrió un error al procesar los archivos.")
