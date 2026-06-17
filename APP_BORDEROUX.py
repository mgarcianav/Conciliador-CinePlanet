from datetime import datetime
import io
import re
import zipfile
import pandas as pd
import streamlit as st


# =====================================================================
# 1. FUNCIONES DE LIMPIEZA Y PROCESAMIENTO (FUENTE A: BORDEROUX)
# =====================================================================
def limpiar_nombre_pelicula(texto_celda):
    texto_celda = str(texto_celda).strip()
    for palabra_clave in [
        " APTA",
        " MAYORES",
        " - ",
        " 1 SEMANA",
        " 2 SEMANA",
        " 3 SEMANA",
    ]:
        if palabra_clave in texto_celda.upper():
            match = re.search(re.escape(palabra_clave), texto_celda, re.IGNORECASE)
            if match:
                texto_celda = texto_celda[: match.start()].strip()

    regex_formatos = (
        r"\s*\((2D|3D|DU|DO|SU|XD|IMAX|4X|PR|DX|XT|IX|SU-4X|DU-XD|DU-2D)\)"
    )
    texto_limpio = re.sub(regex_formatos, "", texto_celda, flags=re.IGNORECASE)
    match_idiomas = re.search(r"([^\(]+)\s*\(([^\)]+)\)", texto_limpio)

    if match_idiomas:
        titulo_en = match_idiomas.group(1).strip()
        titulo_es = match_idiomas.group(2).strip()
    else:
        titulo_en = texto_limpio.strip()
        titulo_es = titulo_en

    return (
        titulo_es.replace("(", "").replace(")", "").strip(),
        titulo_en.replace("(", "").replace(")", "").strip(),
    )


def limpiar_monto_numerico(valor):
    if pd.isna(valor):
        return 0
    val_str = (
        str(valor)
        .upper()
        .replace("S/.", "")
        .replace("$", "")
        .replace(",", "")
        .strip()
    )
    if val_str == "-" or val_str == "" or "TOTAL" in val_str:
        return 0
    try:
        return int(float(val_str))
    except ValueError:
        return 0


def procesar_excel_dinamico(file_bytes, nombre_archivo):
    df_raw = pd.read_excel(file_bytes, header=None)
    data_rows = []
    current_movie_es = None
    current_movie_en = None

    idx_admis = 17
    idx_gbo = 18

    for idx, row in df_raw.iterrows():
        if pd.isna(row.iloc[0]):
            continue

        row_str = str(row.iloc[0]).strip()

        if (
            ("APTA" in row_str.upper() or "MAYORES" in row_str.upper())
            and "CINES" not in row_str.upper()
            and "TOTAL" not in row_str.upper()
        ):
            current_movie_es, current_movie_en = limpiar_nombre_pelicula(
                row_str
            )
            continue

        if (
            row_str.upper().startswith("CINEPLANET")
            and "TOTAL" not in row_str.upper()
            and current_movie_en
        ):
            palabras = row_str.split(" ")
            distribuidor = palabras[0]

            sala_str = row_str[-2:].strip()
            try:
                sala = int(sala_str)
                cine_complejo = row_str[len(distribuidor) : -2].strip()
            except ValueError:
                sala = None
                cine_complejo = row_str[len(distribuidor) :].strip()

            if "CUZCO" in cine_complejo.upper():
                cine_complejo = "CUSCO"
            elif cine_complejo.upper() == "HUANCAYO":
                cine_complejo = "HUANCAYO REAL PLAZA"

            try:
                sem_admis = limpiar_monto_numerico(row.iloc[idx_admis])
                sem_gbo = limpiar_monto_numerico(row.iloc[idx_gbo])
            except IndexError:
                sem_admis = 0
                sem_gbo = 0

            data_rows.append(
                {
                    "Pelicula": current_movie_es.upper(),
                    "Cine": cine_complejo.upper(),
                    "Sala": str(sala) if sala else "NE",
                    "Admisiones_Borderoux": sem_admis,
                    "GBO_Borderoux": sem_gbo,
                }
            )

    return pd.DataFrame(data_rows)


def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", nombre)


def concatenar_llave(df):
    return (
        df["Pelicula"].astype(str).str.strip().str.upper()
        + "_"
        + df["Cine"].astype(str).str.strip().str.upper()
        + "_S"
        + df["Sala"].astype(str).str.strip()
    )


# =====================================================================
# 3. INTERFAZ GRÁFICA DE STREAMLIT (ESTILO DASHBOARD MODULAR)
# =====================================================================
st.set_page_config(
    page_title="Conciliador de Reportes de Cine", layout="wide"
)
st.title("🎬 Analítico y Conciliador de Taquilla")

tab_borderoux, tab_fuente_b, tab_comparativo = st.tabs(
    [
        "📥 1. Fuente A: Borderoux",
        "📥 2. Fuente B: Reportes Consolidados",
        "📊 3. Ver Módulos de Comparativo",
    ]
)

if "df_borderoux" not in st.session_state:
    st.session_state.df_borderoux = None
if "df_fuente_b" not in st.session_state:
    st.session_state.df_fuente_b = None
if "resultados_cruce" not in st.session_state:
    st.session_state.resultados_cruce = None

# --- PESTAÑA 1: FUENTE A ---
with tab_borderoux:
    st.subheader("Carga de archivos Borderoux originales (Cineplanet)")
    archivos_a = st.file_uploader(
        "Sube uno o varios archivos de Borderoux",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uploader_a",
    )

    if archivos_a:
        if st.button("📌 Procesar e indexar Fuente A"):
            listado_dfs = []
            for arc in archivos_a:
                try:
                    df_aux = procesar_excel_dinamico(arc, arc.name)
                    if not df_aux.empty:
                        listado_dfs.append(df_aux)
                except Exception as e:
                    st.error(f"Error en {arc.name}: {e}")

            if listado_dfs:
                df_consolidado_a = pd.concat(listado_dfs, ignore_index=True)
                df_consolidado_a = (
                    df_consolidado_a.groupby(["Pelicula", "Cine", "Sala"])
                    .sum()
                    .reset_index()
                )
                st.session_state.df_borderoux = df_consolidado_a
                st.success(
                    f"¡Fuente A lista! Registros estructurados totales: {len(df_consolidado_a)}"
                )
                st.dataframe(df_consolidado_a, use_container_width=True)

# --- PESTAÑA 2: FUENTE B (MODIFICADO PARA MÚLTIPLES ARCHIVOS) ---
with tab_fuente_b:
    st.subheader("Carga de reportes consolidados externos")
    st.caption("Estructura de columnas soportada: PELÍCULA, CINE, NRO SALA, ADMITS, GROSS TOTAL")

    archivos_b = st.file_uploader(
        "Sube uno o varios archivos Excel de la segunda fuente (Consolidados)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,  # PERMITE CARGAR MÁS DE UN ARCHIVO
        key="uploader_b",
    )

    if archivos_b:
        st.write(f"📁 Archivos Fuente B seleccionados: **{len(archivos_b)}**")

        if st.button("📌 Procesar e indexar Fuente B Masiva"):
            listado_dfs_b = []
            
            for arc_b in archivos_b:
                try:
                    df_b_raw = pd.read_excel(arc_b)
                    df_b_raw.columns = [str(c).strip().upper() for c in df_b_raw.columns]

                    if "PELÍCULA" in df_b_raw.columns and "CINE" in df_b_raw.columns:
                        df_b_final = pd.DataFrame()
                        df_b_final["Pelicula"] = df_b_raw["PELÍCULA"].astype(str).str.strip().str.upper()
                        df_b_final["Cine"] = df_b_raw["CINE"].astype(str).str.strip().str.upper()

                        if "NRO SALA" in df_b_raw.columns:
                            df_b_final["Sala"] = (
                                df_b_raw["NRO SALA"]
                                .astype(str)
                                .str.replace(".0", "", regex=False)
                                .str.strip()
                            )
                        else:
                            df_b_final["Sala"] = "NE"

                        col_admits = [c for c in df_b_raw.columns if "ADMITS" in c]
                        df_b_final["Admisiones_FuenteB"] = (
                            df_b_raw[col_admits[0]].fillna(0).astype(int) if col_admits else 0
                        )

                        col_gross = [c for c in df_b_raw.columns if "GROSS" in c]
                        df_b_final["GBO_FuenteB"] = (
                            df_b_raw[col_gross[0]].fillna(0).astype(int) if col_gross else 0
                        )

                        listado_dfs_b.append(df_b_final)
                    else:
                        st.warning(f"Columnas inválidas en el archivo '{arc_b.name}'. Saltando archivo.")
                except Exception as e:
                    st.error(f"Error procesando el archivo '{arc_b.name}': {e}")

            if listado_dfs_b:
                # Unificar y agrupar todo el bloque de archivos de la Fuente B
                df_consolidado_b = pd.concat(listado_dfs_b, ignore_index=True)
                df_consolidado_b = (
                    df_consolidado_b.groupby(["Pelicula", "Cine", "Sala"])
                    .sum()
                    .reset_index()
                )
                df_consolidado_b["Cine"] = df_consolidado_b["Cine"].str.replace("CUZCO", "CUSCO")

                st.session_state.df_fuente_b = df_consolidado_b
                st.success(
                    f"¡Fuente B consolidada con éxito! Registros estructurados totales: {len(df_consolidado_b)}"
                )
                st.dataframe(df_consolidado_b, use_container_width=True)
            else:
                st.error("No se pudo extraer información válida de ningún archivo cargado en Fuente B.")

# --- PESTAÑA 3: COMPARATIVO MODULAR CON MÉTRICAS Y VISTA PREVIA ---
with tab_comparativo:
    st.subheader("Cruce y Auditoría Automatizada")

    if st.session_state.df_borderoux is not None and st.session_state.df_fuente_b is not None:
        df_a = st.session_state.df_borderoux.copy()
        df_b = st.session_state.df_fuente_b.copy()

        df_a["LLAVE_CRUCE"] = concatenar_llave(df_a)
        df_b["LLAVE_CRUCE"] = concatenar_llave(df_b)

        fecha_actual = datetime.now().strftime("%d%b%Y").upper()

        if st.button("🔍 Correr Análisis de Conciliación"):
            with st.spinner("Comparando registros llave a llave..."):
                todas_peliculas = pd.concat([df_a["Pelicula"], df_b["Pelicula"]]).unique()
                diccionario_resultados = {}

                for pelicula in todas_peliculas:
                    sub_a = df_a[df_a["Pelicula"] == pelicula]
                    sub_b = df_b[df_b["Pelicula"] == pelicula]

                    df_merge = pd.merge(
                        sub_a,
                        sub_b,
                        on="LLAVE_CRUCE",
                        how="outer",
                        suffixes=("_A", "_B"),
                    )

                    df_merge["Pelicula"] = df_merge["Pelicula_A"].fillna(df_merge["Pelicula_B"])
                    df_merge["Cine"] = df_merge["Cine_A"].fillna(df_merge["Cine_B"])
                    df_merge["Sala"] = df_merge["Sala_A"].fillna(df_merge["Sala_B"])

                    df_merge["Admisiones_Borderoux"] = df_merge["Admisiones_Borderoux"].fillna(0)
                    df_merge["GBO_Borderoux"] = df_merge["GBO_Borderoux"].fillna(0)
                    df_merge["Admisiones_FuenteB"] = df_merge["Admisiones_FuenteB"].fillna(0)
                    df_merge["GBO_FuenteB"] = df_merge["GBO_FuenteB"].fillna(0)

                    # MATCH PERFECTO
                    cond_match = (
                        (df_merge["Pelicula_A"].notna())
                        & (df_merge["Pelicula_B"].notna())
                        & (df_merge["Admisiones_Borderoux"] == df_merge["Admisiones_FuenteB"])
                        & (df_merge["GBO_Borderoux"] == df_merge["GBO_FuenteB"])
                    )
                    df_match = df_merge[cond_match][["Cine", "Sala", "Admisiones_Borderoux", "GBO_Borderoux"]]

                    # DIFERENCIAS NUMÉRICAS
                    cond_dif = (
                        (df_merge["Pelicula_A"].notna())
                        & (df_merge["Pelicula_B"].notna())
                        & (
                            (df_merge["Admisiones_Borderoux"] != df_merge["Admisiones_FuenteB"])
                            | (df_merge["GBO_Borderoux"] != df_merge["GBO_FuenteB"])
                        )
                    )
                    df_diferencias = df_merge[cond_dif].copy()
                    if not df_diferencias.empty:
                        df_diferencias["Dif_Admisiones"] = (
                            df_diferencias["Admisiones_Borderoux"] - df_diferencias["Admisiones_FuenteB"]
                        )
                        df_diferencias["Dif_GBO"] = (
                            df_diferencias["GBO_Borderoux"] - df_diferencias["GBO_FuenteB"]
                        )
                        df_diferencias = df_diferencias[
                            [
                                "Cine",
                                "Sala",
                                "Admisiones_Borderoux",
                                "Admisiones_FuenteB",
                                "Dif_Admisiones",
                                "GBO_Borderoux",
                                "GBO_FuenteB",
                                "Dif_GBO",
                            ]
                        ]

                    # OMISIONES
                    df_solo_borderoux = df_merge[df_merge["Pelicula_B"].isna()][
                        ["Cine", "Sala", "Admisiones_Borderoux", "GBO_Borderoux"]
                    ]
                    df_solo_fuenteb = df_merge[df_merge["Pelicula_A"].isna()][
                        ["Cine", "Sala", "Admisiones_FuenteB", "GBO_FuenteB"]
                    ]

                    diccionario_resultados[pelicula] = {
                        "match": df_match,
                        "diferencias": df_diferencias,
                        "solo_borderoux": df_solo_borderoux,
                        "solo_fuente_b": df_solo_fuenteb,
                    }

                st.session_state.resultados_cruce = diccionario_resultados
                st.success("¡Datos cruzados en memoria listos para previsualizar!")

        # --- SECCIÓN VISUAL DEL DASHBOARD ---
        if st.session_state.resultados_cruce is not None:
            st.write("---")

            lista_peliculas = list(st.session_state.resultados_cruce.keys())
            pelicula_seleccionada = st.selectbox(
                "🎬 Selecciona la Película para Auditar su Grilla de Vista Previa:",
                lista_peliculas,
            )

            res = st.session_state.resultados_cruce[pelicula_seleccionada]

            # MÓDULOS DE KPI
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(label="✨ Match Perfecto", value=f"{len(res['match'])} cines")
            c2.metric(
                label="⚠️ Diferencias en Montos",
                value=f"{len(res['diferencias'])} cines",
                delta="Alerta" if len(res["diferencias"]) > 0 else "Ok",
                delta_color="inverse",
            )
            c3.metric(label="❌ Faltan en Fuente B", value=f"{len(res['solo_borderoux'])} salas")
            c4.metric(label="🚨 Faltan en Borderoux", value=f"{len(res['solo_fuente_b'])} salas")

            # GRILLA ANALÍTICA INTERACTIVA (VISTA PREVIA)
            st.subheader(f"📊 Vista Previa de Datos: {pelicula_seleccionada}")
            v1, v2, v3, v4 = st.tabs(
                [
                    "✅ Coincidencias",
                    "🔍 Diferencias Numéricas",
                    "📌 Solo en Borderoux",
                    "📌 Solo en Reporte B",
                ]
            )

            with v1:
                st.dataframe(res["match"], use_container_width=True)
            with v2:
                st.dataframe(res["diferencias"], use_container_width=True)
            with v3:
                st.dataframe(res["solo_borderoux"], use_container_width=True)
            with v4:
                st.dataframe(res["solo_fuente_b"], use_container_width=True)

            # DESCARGA DEL ZIP
            st.write("---")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as nuevo_zip:
                for pel, datos in st.session_state.resultados_cruce.items():
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        datos["match"].to_excel(writer, sheet_name="MATCH PERFECTO", index=False)
                        if not datos["diferencias"].empty:
                            datos["diferencias"].to_excel(writer, sheet_name="DIFERENCIAS NUMÉRICAS", index=False)
                        if not datos["solo_borderoux"].empty:
                            datos["solo_borderoux"].to_excel(writer, sheet_name="SOLO EN BORDEROUX", index=False)
                        if not datos["solo_fuente_b"].empty:
                            datos["solo_fuente_b"].to_excel(writer, sheet_name="SOLO EN FUENTE B", index=False)

                    n_limpio = limpiar_nombre_archivo(pel).upper()
                    nuevo_zip.writestr(
                        f"BORDEROUX_COMP_{n_limpio}_{fecha_actual}.xlsx",
                        excel_buffer.getvalue(),
                    )

            st.download_button(
                label="📥 Descargar todos los comparativos auditados (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"CONCILIACION_TAQUILLA_{fecha_actual}.zip",
                mime="application/zip",
            )
    else:
        st.warning(
            "Carga la Fuente A y la Fuente B en sus respectivas pestañas para activar la conciliación modular."
        )