import pandas as pd
from io import BytesIO

def procesar_oa2(file_pasado, file_actual):
    """
    Procesa dos archivos Excel (pasado y actual) según el enfoque metodológico de análisis de expedientes.
    Devuelve un BytesIO con el Excel listo para descargar y un diccionario con las comparaciones por prefijo.
    """
    try:
        # Leer archivos Excel
        df_pasado = pd.read_excel(file_pasado, dtype=str)
        df_actual = pd.read_excel(file_actual, dtype=str)

        # Asegurar que MONTO es numérico
        df_pasado["MONTO"] = pd.to_numeric(df_pasado.get("MONTO", 0), errors="coerce").fillna(0)
        df_actual["MONTO"] = pd.to_numeric(df_actual.get("MONTO", 0), errors="coerce").fillna(0)

        # Crear columnas datounico y cuenta
        def crear_tabla(df):
            df["datounico"] = (
                df["EXPEDIENTE / CASO"].astype(str) + "-" +
                df["NUM_DOC_DEMANDANTE"].astype(str) + "-" +
                df["DEMANDANTE_NOMBRE"].astype(str)
            )
            df["cuenta"] = df["MAYOR"].astype(str) + "-" + df["SUB_CTA"].astype(str)
            df_agrupado = df.groupby(["datounico", "cuenta"], as_index=False)["MONTO"].sum()
            return df_agrupado

        df_pasado = crear_tabla(df_pasado)
        df_actual = crear_tabla(df_actual)

        # Comparación por prefijo de cuenta
        def comparar_por_prefijo(df_pasado, df_actual, prefijo):
            resultados = []
            pasado_pref = df_pasado[df_pasado["cuenta"].str.startswith(prefijo)]
            actual_pref = df_actual[df_actual["cuenta"].str.startswith(prefijo)]

            for _, row in pasado_pref.iterrows():
                datounico = row["datounico"]
                cuenta_pasado = row["cuenta"]
                monto_pasado = row["MONTO"]

                match = actual_pref[actual_pref["datounico"] == datounico]

                if not match.empty:
                    cuentas_actuales = match["cuenta"].unique()
                    if cuenta_pasado in cuentas_actuales:
                        monto_actual = match.loc[match["cuenta"] == cuenta_pasado, "MONTO"].sum()
                        diferencia = monto_actual - monto_pasado
                        resultados.append([datounico, cuenta_pasado, cuenta_pasado, monto_pasado, monto_actual, diferencia, "Misma cuenta"])
                    else:
                        monto_total = match["MONTO"].sum()
                        resultados.append([datounico, cuenta_pasado, ", ".join(cuentas_actuales), monto_pasado, monto_total, None, "Cuenta diferente"])
                else:
                    resultados.append([datounico, cuenta_pasado, "-", monto_pasado, 0, -monto_pasado, "Solo en pasado"])

            for _, row in actual_pref.iterrows():
                datounico = row["datounico"]
                cuenta_actual = row["cuenta"]
                monto_actual = row["MONTO"]
                if datounico not in pasado_pref["datounico"].values:
                    resultados.append([datounico, "-", cuenta_actual, 0, monto_actual, monto_actual, "Solo en actual"])

            return pd.DataFrame(resultados, columns=["datounico", "Cuenta_Pasado", "Cuenta_Actual",
                                                     "MONTO_PASADO", "MONTO_ACTUAL", "Diferencia", "Resultado"])

        prefijos = ["1202", "9110", "2401", "2103"]
        comparaciones = {f"Comparación {p}": comparar_por_prefijo(df_pasado, df_actual, p) for p in prefijos}

        # Exportar a Excel en memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_pasado.to_excel(writer, index=False, sheet_name="PASADO")
            df_actual.to_excel(writer, index=False, sheet_name="ACTUAL")
            for nombre, df_comp in comparaciones.items():
                df_comp.to_excel(writer, index=False, sheet_name=nombre)
        output.seek(0)

        return output, comparaciones

    except Exception as e:
        print(f"Error al procesar OA-2: {e}")
        return None, None
