import pandas as pd
from io import BytesIO

def procesar_oa2(file_pasado, file_actual):
    """
    Procesa dos archivos Excel (pasado y actual) según el enfoque metodológico de análisis de expedientes.
    Se compara la columna completa 'cuenta', incluyendo subcuenta analítica.
    Devuelve un BytesIO con el Excel listo para descargar y un DataFrame con la comparación.
    """
    try:
        # Leer archivos Excel
        df_pasado = pd.read_excel(file_pasado, dtype=str)
        df_actual = pd.read_excel(file_actual, dtype=str)

        # MONTO numérico
        df_pasado["MONTO"] = pd.to_numeric(df_pasado.get("MONTO", 0), errors="coerce").fillna(0)
        df_actual["MONTO"] = pd.to_numeric(df_actual.get("MONTO", 0), errors="coerce").fillna(0)

        # Crear datounico y cuenta
        def crear_tabla(df):
            df["datounico"] = (
                df["EXPEDIENTE / CASO"].astype(str) + "-" +
                df["NUM_DOC_DEMANDANTE"].astype(str) + "-" +
                df["DEMANDANTE_NOMBRE"].astype(str)
            )
            df["cuenta"] = df["MAYOR"].astype(str) + "-" + df["SUB_CTA"].astype(str)
            # Agrupar por datounico y cuenta, sumando MONTO
            return df.groupby(["datounico", "cuenta"], as_index=False)["MONTO"].sum()

        df_pasado = crear_tabla(df_pasado)
        df_actual = crear_tabla(df_actual)

        # Comparación por cuenta completa
        def comparar_por_cuenta(df_pasado, df_actual):
            resultados = []
            for _, row in df_pasado.iterrows():
                datounico = row["datounico"]
                cuenta_pasado = row["cuenta"]
                monto_pasado = row["MONTO"]

                match = df_actual[(df_actual["datounico"] == datounico) & (df_actual["cuenta"] == cuenta_pasado)]

                if not match.empty:
                    monto_actual = match["MONTO"].sum()
                    diferencia = monto_actual - monto_pasado
                    resultados.append([datounico, cuenta_pasado, cuenta_pasado, monto_pasado, monto_actual, diferencia, "Misma cuenta"])
                else:
                    match_diff = df_actual[df_actual["datounico"] == datounico]
                    if not match_diff.empty:
                        cuentas_actuales = match_diff["cuenta"].unique()
                        monto_total = match_diff["MONTO"].sum()
                        resultados.append([datounico, cuenta_pasado, ", ".join(cuentas_actuales), monto_pasado, monto_total, None, "Cuenta diferente"])
                    else:
                        resultados.append([datounico, cuenta_pasado, "-", monto_pasado, 0, -monto_pasado, "Solo en pasado"])

            # Filas solo en actual
            for _, row in df_actual.iterrows():
                datounico = row["datounico"]
                cuenta_actual = row["cuenta"]
                monto_actual = row["MONTO"]
                if datounico not in df_pasado["datounico"].values:
                    resultados.append([datounico, "-", cuenta_actual, 0, monto_actual, monto_actual, "Solo en actual"])

            return pd.DataFrame(resultados, columns=["datounico", "Cuenta_Pasado", "Cuenta_Actual", "MONTO_PASADO", "MONTO_ACTUAL", "Diferencia", "Resultado"])

        comparacion_total = comparar_por_cuenta(df_pasado, df_actual)

        # Exportar a Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_pasado.to_excel(writer, index=False, sheet_name="PASADO")
            df_actual.to_excel(writer, index=False, sheet_name="ACTUAL")
            comparacion_total.to_excel(writer, index=False, sheet_name="Comparación_Cuenta")

        output.seek(0)
        return output, comparacion_total

    except Exception as e:
        print(f"Error al procesar OA-2: {e}")
        return None, None
