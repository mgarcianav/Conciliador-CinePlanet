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

    fechas_mapeo = {
        "2026-07-09": 2,
        "2026-07-10": 4,
        "2026-07-11": 6,
        "2026-07-12": 8,
        "2026-07-13": 12,
        "2026-07-14": 14,
        "2026-07-15": 16
    }

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

            gbo_diario = {}
            for f_str, col_idx in fechas_mapeo.items():
                try:
                    gbo_diario[f_str] = limpiar_monto_numerico(row.iloc[col_idx])
                except:
                    gbo_diario[f_str] = 0

            data_rows.append(
                {
                    "Pelicula": current_movie_es.upper(),
                    "Cine": cine_complejo.upper(),
                    "Sala": str(sala) if sala else "NE",
                    "Admisiones_Borderoux": sem_admis,
                    "GBO_Borderoux": sem_gbo,
                    "GBO_Diario_A": gbo_diario
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
if "df_fuente_b_detalle" not in st.session_state:
    st.session_state.df_fuente_b_detalle = None
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
                
                def combinar_gbo_diarios(series):
                    consolidado = {}
                    for d in series:
                        if isinstance(d, dict):
                            for k, v in d.items():
                                consolidado[k] = consolidado.get(k, 0) + v
                    return consolidado

                df_consolidado_a = df_consolidado_a.groupby(["Pelicula", "Cine", "Sala"]).agg({
                    "Admisiones_Borderoux": "sum",
                    "GBO_Borderoux": "sum",
                    "GBO_Diario_A": combinar_gbo_diarios
                }).reset_index()

                st.session_state.df_borderoux = df_consolidado_a
                st.success(
                    f"¡Fuente A lista! Registros estructurados totales: {len(df_consolidado_a)}"
                )
                st.dataframe(df_consolidado_a.drop(columns=["GBO_Diario_A"], errors="ignore"), use_container_width=True)

# --- PESTAÑA 2: FUENTE B ---
with tab_fuente_b:
    st.subheader("Carga de reportes consolidados externos")
    st.caption("Debe contener las pestañas 'CONSOLIDADO' y 'DETALLE'. Estructura soportada: PELÍCULA, CINE, NRO SALA, ADMITS, GROSS TOTAL")

    archivos_b = st.file_uploader(
        "Sube uno o varios archivos Excel de la segunda fuente (Consolidados)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="uploader_b",
    )

    if archivos_b:
        st.write(f"📁 Archivos Fuente B seleccionados: **{len(archivos_b)}**")

        if st.button("📌 Procesar e indexar Fuente B Masiva"):
            listado_dfs_b = []
            listado_dfs_b_detalle = []
            
            for arc_b in archivos_b:
                try:
                    xls = pd.ExcelFile(arc_b)
                    sheet_consolidado = "CONSOLIDADO" if "CONSOLIDADO" in [s.upper() for s in xls.sheet_names] else xls.sheet_names[0]
                    
                    df_b_raw = pd.read_excel(xls, sheet_name=sheet_consolidado)
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
                    
                    sheet_detalle = next((s for s in xls.sheet_names if "DETALLE" in s.upper()), None)
                    if sheet_detalle:
                        df_b_det = pd.read_excel(xls, sheet_name=sheet_detalle)
                        df_b_det.columns = [str(c).strip().upper() for c in df_b_det.columns]
                        
                        if "PELÍCULA" in df_b_det.columns and "CINE" in df_b_det.columns:
                            df_b_det["FECHA_PARSED"] = pd.to_datetime(df_b_det["FECHA"].astype(str).str.strip(), format="%d/%m/%Y", errors='coerce')
                            df_b_det["Pelicula"] = df_b_det["PELÍCULA"].astype(str).str.strip().str.upper()
                            df_b_det["Cine"] = df_b_det["CINE"].astype(str).str.strip().str.upper().str.replace("CUZCO", "CUSCO")
                            df_b_det["Sala"] = df_b_det["NRO SALA"].astype(str).str.replace(".0", "", regex=False).str.strip()
                            df_b_det["LLAVE_CRUCE"] = df_b_det["Pelicula"] + "_" + df_b_det["Cine"] + "_S" + df_b_det["Sala"]
                            
                            listado_dfs_b_detalle.append(df_b_det)

                except Exception as e:
                    st.error(f"Error procesando el archivo '{arc_b.name}': {e}")

            if listado_dfs_b:
                df_consolidado_b = pd.concat(listado_dfs_b, ignore_index=True)
                df_consolidado_b = (
                    df_consolidado_b.groupby(["Pelicula", "Cine", "Sala"])
                    .sum()
                    .reset_index()
                )
                df_consolidado_b["Cine"] = df_consolidado_b["Cine"].str.replace("CUZCO", "CUSCO")

                st.session_state.df_fuente_b = df_consolidado_b
                
                if listado_dfs_b_detalle:
                    st.session_state.df_fuente_b_detalle = pd.concat(listado_dfs_b_detalle, ignore_index=True)

                st.success(
                    f"¡Fuente B consolidada! Registros acumulados: {len(df_consolidado_b)} (Detalles diarios indexados: {len(st.session_state.df_fuente_b_detalle) if st.session_state.df_fuente_b_detalle is not None else 0})"
                )
                st.dataframe(df_consolidado_b, use_container_width=True)

# --- PESTAÑA 3: COMPARATIVO MODULAR ---
with tab_comparativo:
    st.subheader("Cruce y Auditoría Automatizada con Análisis Diario")

    if st.session_state.df_borderoux is not None and st.session_state.df_fuente_b is not None:
        df_a = st.session_state.df_borderoux.copy()
        df_b = st.session_state.df_fuente_b.copy()

        df_a["LLAVE_CRUCE"] = concatenar_llave(df_a)
        df_b["LLAVE_CRUCE"] = concatenar_llave(df_b)

        fecha_actual = datetime.now().strftime("%d%b%Y").upper()

        if st.button("🔍 Correr Análisis de Conciliación"):
            with st.spinner("Comparando registros llave a llave y detectando diferencias por fecha..."):
                todas_peliculas = pd.concat([df_a["Pelicula"], df_b["Pelicula"]]).unique()
                diccionario_resultados = {}

                df_b_detalle = st.session_state.df_fuente_b_detalle

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

                    df_merge["Admisiones_Borderoux"] = df_merge["Admisiones_Borderoux"].fillna(0).astype(int)
                    df_merge["GBO_Borderoux"] = df_merge["GBO_Borderoux"].fillna(0).astype(int)
                    df_merge["Admisiones_FuenteB"] = df_merge["Admisiones_FuenteB"].fillna(0).astype(int)
                    df_merge["GBO_FuenteB"] = df_merge["GBO_FuenteB"].fillna(0).astype(int)

                    # MATCH PERFECTO
                    cond_match = (
                        (df_merge["Pelicula_A"].notna())
                        & (df_merge["Pelicula_B"].notna())
                        & (df_merge["Admisiones_Borderoux"] == df_merge["Admisiones_FuenteB"])
                        & (df_merge["GBO_Borderoux"] == df_merge["GBO_FuenteB"])
                    )
                    df_match = df_merge[cond_match][["Cine", "Sala", "Admisiones_Borderoux", "GBO_Borderoux"]]

                    # DIFERENCIAS NUMÉRICAS Y ANÁLISIS DÍA A DÍA
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
                        ).astype(int)
                        df_diferencias["Dif_GBO"] = (
                            df_diferencias["GBO_Borderoux"] - df_diferencias["GBO_FuenteB"]
                        ).astype(int)
                        
                        observaciones_diarias = []
                        for idx_dif, row_dif in df_diferencias.iterrows():
                            llave = row_dif["LLAVE_CRUCE"]
                            gbo_diario_a = row_dif["GBO_Diario_A"] if isinstance(row_dif["GBO_Diario_A"], dict) else {}
                            
                            fechas_con_diferencia = []
                            
                            if df_b_detalle is not None and not df_b_detalle.empty:
                                detalle_b = df_b_detalle[df_b_detalle["LLAVE_CRUCE"] == llave].copy()
                                detalle_b["FECHA_STR"] = detalle_b["FECHA_PARSED"].dt.strftime('%Y-%m-%d')
                                
                                for f_str, gbo_a in gbo_diario_a.items():
                                    reg_b = detalle_b[detalle_b["FECHA_STR"] == f_str]
                                    
                                    gbo_b = 0
                                    if not reg_b.empty:
                                        try:
                                            gbo_b = int(float(str(reg_b["GROSS TOTAL"].values[0]).replace(",","").strip()))
                                        except ValueError:
                                            gbo_b = 0
                                    
                                    if abs(gbo_a - gbo_b) > 1:
                                        f_label = pd.to_datetime(f_str).strftime('%d-%b')
                                        fechas_con_diferencia.append(f"{f_label} (A: S/. {gbo_a:,} vs B: S/. {gbo_b:,})")
                                        
                            if fechas_con_diferencia:
                                obs_text = "Diferencias en fechas: " + ", ".join(fechas_con_diferencia)
                            else:
                                # MEJORA AQUÍ: Si no hay cruce en DETALLE pero los totales difieren, extrae los días con taquilla activa en Borderoux para mostrarlos.
                                dias_activos = [pd.to_datetime(f).strftime('%d-%b') for f, v in gbo_diario_a.items() if v > 0]
                                f_rango = ", ".join(dias_activos) if dias_activos else "Rango cargado"
                                obs_text = f"Diferencia en Fechas [{f_rango}] -> Acumulado Total difiere (Bdx: S/. {row_dif['GBO_Borderoux']:,} vs Fuente B: S/. {row_dif['GBO_FuenteB']:,})."
                            
                            observaciones_diarias.append(obs_text)
                        
                        df_diferencias["OBSERVACIONES"] = observaciones_diarias
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
                                "OBSERVACIONES"
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
                st.success("¡Datos cruzados con auditoría de desglose diario listos!")

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