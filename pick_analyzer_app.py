import streamlit as st
import pandas as pd
import io
import os

# ============================================================
# CONFIGURACIÓN
# ============================================================
COLUMNA_PICKLIST  = "Picklist Code"
COLUMNA_ASSOCIATE = "Associate"
COLUMNA_BAGS      = "Bags"
COLUMNA_OVS       = "OVs"
COLUMNA_DURATION  = "Duration"


# ============================================================
# FUNCIONES DE PROCESAMIENTO
# (Son exactamente las mismas que en el script original)
# ============================================================
def procesar_datos(df):

    mask_subpick = df[COLUMNA_PICKLIST].astype(str).str.contains("#", na=False)
    df_filtrado = df[mask_subpick].copy()

    mask_individual = ~df_filtrado[COLUMNA_ASSOCIATE].astype(str).str.contains(",", na=False)
    df_filtrado = df_filtrado[mask_individual].copy()

    df_filtrado = df_filtrado[df_filtrado[COLUMNA_ASSOCIATE].notna()]
    df_filtrado = df_filtrado[df_filtrado[COLUMNA_ASSOCIATE].astype(str).str.strip() != ""]

    def duration_a_segundos(valor):
        try:
            partes = str(valor).split(":")
            horas    = int(partes[0])
            minutos  = int(partes[1])
            segundos = int(partes[2].split(".")[0])
            return horas * 3600 + minutos * 60 + segundos
        except Exception:
            try:
                return float(valor) * 86400
            except Exception:
                return None

    df_filtrado["Duration_seg"] = df_filtrado[COLUMNA_DURATION].apply(duration_a_segundos)
    df_filtrado[COLUMNA_BAGS]   = pd.to_numeric(df_filtrado[COLUMNA_BAGS], errors="coerce").fillna(0)
    df_filtrado[COLUMNA_OVS]    = pd.to_numeric(df_filtrado[COLUMNA_OVS],  errors="coerce").fillna(0)

    tabla = df_filtrado.groupby(COLUMNA_ASSOCIATE).agg(
        Num_Picklists    = (COLUMNA_PICKLIST, "count"),
        Total_Bags       = (COLUMNA_BAGS,     "sum"),
        Total_OVs        = (COLUMNA_OVS,      "sum"),
        Avg_Duration_seg = ("Duration_seg",   "mean")
    ).reset_index()

    def segundos_a_mmss(seg):
        try:
            seg      = int(seg)
            minutos  = seg // 60
            segundos = seg % 60
            return f"{minutos:02d}:{segundos:02d}"
        except Exception:
            return "-"

    tabla["Avg_Duration"] = tabla["Avg_Duration_seg"].apply(segundos_a_mmss)
    tabla = tabla.drop(columns=["Avg_Duration_seg"])
    tabla = tabla.sort_values("Num_Picklists", ascending=False).reset_index(drop=True)

    return tabla


def convertir_a_excel(tabla):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabla.to_excel(writer, sheet_name="Tabla Dinámica", index=False)
        ws = writer.sheets["Tabla Dinámica"]
        for col in ws.columns:
            max_ancho = max(len(str(c.value)) for c in col if c.value is not None)
            ws.column_dimensions[col[0].column_letter].width = max_ancho + 4
    buffer.seek(0)
    return buffer


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================
st.set_page_config(
    page_title="PICK Analyzer - DCZ4",
    page_icon="📦",
    layout="wide"
)

st.title("📦 PICK Analyzer - DCZ4")
st.markdown("Sube el archivo PICK del día y obtén automáticamente la tabla de rendimiento por asociado.")

st.divider()

# --- Uploader
archivo = st.file_uploader(
    "Selecciona el archivo PICK (.xlsx o .csv)",
    type=["xlsx", "xls", "csv"]
)

if archivo is not None:

    # Cargar según extensión
    nombre = archivo.name
    extension = os.path.splitext(nombre)[1].lower()

    try:
        with st.spinner("Procesando datos..."):
            if extension in [".xlsx", ".xls"]:
                df = pd.read_excel(archivo, sheet_name=0, header=1)
            else:
                df = pd.read_csv(archivo)

            tabla = procesar_datos(df)

        st.success(f"✅ Procesado correctamente. {len(tabla)} asociados encontrados.")

        st.divider()

        # --- Métricas resumen en la parte superior
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Asociados",  len(tabla))
        col2.metric("Total Bags",        int(tabla["Total_Bags"].sum()))
        col3.metric("Total OVs",         int(tabla["Total_OVs"].sum()))
        col4.metric("Total Picklists",   int(tabla["Num_Picklists"].sum()))

        st.divider()

        # --- Tabla de resultados
        st.subheader("📊 Rendimiento por Asociado")
        st.dataframe(
            tabla,
            use_container_width=True,
            hide_index=True,
            column_config={
                COLUMNA_ASSOCIATE: st.column_config.TextColumn("Asociado"),
                "Num_Picklists":   st.column_config.NumberColumn("Picklists",      format="%d"),
                "Total_Bags":      st.column_config.NumberColumn("Total Bags",     format="%d"),
                "Total_OVs":       st.column_config.NumberColumn("Total OVs",      format="%d"),
                "Avg_Duration":    st.column_config.TextColumn("Duración Media"),
            }
        )

        st.divider()

        # --- Gráficos
        st.subheader("📈 Visualizaciones")

        tab1, tab2, tab3 = st.tabs(["Picklists", "Bags", "OVs"])

        with tab1:
            st.bar_chart(
                tabla.set_index(COLUMNA_ASSOCIATE)["Num_Picklists"],
                color="#4F8BF9"
            )
        with tab2:
            st.bar_chart(
                tabla.set_index(COLUMNA_ASSOCIATE)["Total_Bags"],
                color="#21C55D"
            )
        with tab3:
            st.bar_chart(
                tabla.set_index(COLUMNA_ASSOCIATE)["Total_OVs"],
                color="#F97316"
            )

        st.divider()

        # --- Botón de descarga
        excel_buffer = convertir_a_excel(tabla)
        nombre_salida = nombre.replace(".xlsx", "").replace(".csv", "") + "_RESULTADO.xlsx"

        st.download_button(
            label="⬇️ Descargar resultado en Excel",
            data=excel_buffer,
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        st.info("Comprueba que el archivo tiene el formato correcto.")

else:
    st.info("👆 Sube un archivo para comenzar.")
