import streamlit as st
import pandas as pd
import importlib.util
import tempfile
import os
import ast
 # --- Funciones de procesamiento y pivotaje ---
import io
def procesar_excel(df):
    # Procesamiento tipo Procesar.ipynb: limpieza, renombrado, reemplazo de valores
    nuevos_encabezados = [str(columna).split('-')[-1].strip() if '-' in str(columna) else str(columna) for columna in df.columns]
    df.columns = nuevos_encabezados
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    reemplazos = {
        "5 (Supera las expectativas)":5, "5 - Muy satisfecho":5, "Muy satisfecho":5, "5. Supera notablemente las expectativas":5,
        "Supera notablemente las expectativas":5, "5 - Supera notablemente las expectativas":5, "5-Supera notablemente las expectativas":5,
        "Supera las expectativas / Exceeds expectations":5, "5. Supera notablemente mis expectativas":5, "5.Muy por encima de sus expectativas":5,
        "5- Supera notablemente las expectativas":5, "4":4, "4 - Satisfecho":4, "Satisfecho":4, "Supera las expectativas / Exceeds expectations4":4,
        "Supera las expectativas / Exceeds expectations\n4":4, "3":3, "Cumple las expectativas / Meets expectations3":3,
        "Cumple las expectativas / Meets expectations\n3":3, "Ni satisfecho - ni insatisfecho":3, "3 - Ni satisfecho - ni insatisfecho":3,
        "Cumple las expectativas / Meets expectations":3, "2":2, "Por debajo de las expectativas / Below expectations\n2":2,
        "Muy insatisfecho":2, "2 - Insatisfecho":2, "Insatisfecho":2, "1 (Muy por debajo de las expectativas)":1, "1 - Muy insatisfecho":1,
        "Muy por debajo de las expectativas":1, "1 - Muy por Debajo de las Expectativas":1, "1. Muy por debajo de las expectativas":1,
        "1. Muy por debajo de mis expectativas":1, "1.Muy por debajo de sus expectativas":1, "Muy por debajo de las expectativas / Far below expectations":1
    }
    df = df.replace(reemplazos)
    return df

def pivotear_excel(df):
    # Procesamiento tipo pivotaje 1.ipynb: pivotar según columnas detectadas
    def detectar_columnas_caso(df):
        casos=[]
        for columns in df.columns:
            if 'caso' in columns.strip().lower():
                casos.append(columns)
        return casos
    def detectar_columnas_respuesta(df):
        for columns in df.columns:
            if 'respuesta' in columns.strip().lower():
                return columns
    def detectar_column_pregunta(df):
        for columns in df.columns:
            if 'pregunta' in columns.strip().lower():
                return columns
    iindexc=detectar_columnas_caso(df)
    icolumnsc=detectar_column_pregunta(df)
    ivaluesc=detectar_columnas_respuesta(df)
    if iindexc and icolumnsc and ivaluesc:
        data = df.pivot_table(index=iindexc, columns=icolumnsc, values=ivaluesc, aggfunc='first').reset_index()
        return data
    else:
        return df

# Diccionario de oficinas vinculado a un único script común y procesos asociados
diccionario_oficinas = {
    "Dirección de Tecnología Informática y Comunicaciones": {
        "script": "Generararchivoexcel_mantenimiento_tic_trimestre",
        "procesos": ["Sistemas de Información desarrollo, mantenimiento y soporte a usuarios"]
    },
    "Operaciones Tic": {
        "script": "Generararchivoexcel_mantenimiento_tic_CSU",
        "procesos": ["Soporte de Servicios TIC"]
    },
    "Dirección de servicios Administrativos": {
        "script": "Generararchivoexcel_mantenimientoDSA",
        "procesos": ["Mantenimiento DSA"]
    },
    "Admisiones": {
        "script": "Generararchivoexcel_admisiones_posgrado",
        "procesos": ["Satisfacción con respecto al servicio recibido durante el proceso de admisión a posgrado"]
    },
    "Financiamiento Empresarial": {
        "script": "Generararchivoexcel_Financiamiento_Empresarial",
        "procesos": ["Servicio prestado en facturación a través de distintas plataformas"]
    },
    "Dirección Financiera": {
        "script": "Generararchivoexcel_registro_provedores",
        "procesos": ["Registro de Proveedores"]
    },
    "Sección de Compras": {
        "script": "Generararchivoexcel_Adquisicion_bienes",
        "procesos": ["Adquisición de bienes"]
    },
    "Laboratorio de Geotecnia y Materiales de Construcción": {
        "script": "Generararchivoexcel_laboratorio_geotecnia",
        "procesos": ["Servicios del Laboratorio de Geotecnia y Materiales de Construcción"]
    },
    "Departamento de Registro": {
        "script": "Generararchivoexcel_coordinadores",
        "procesos": ["Informe de calidad coordinadores de Pregrado"]
    },
    "Departamenro de Ingeniería Mecánica": {
        "script": "Generararchivoexcel_laboratori_cimm",
        "procesos": ["CIMM"]
    },
    "Tesorería": {
        "script": "Generararchivoexcel_Tesoreria",
        "procesos": ["Recaudo Web - Zona Pago", "Recaudo Web - Place to Pay", "Presencial Caja"]
    },
    "Oficina de Registro": {
        "script": "Generararchivoexcel_oficinaregistro_grado",
        "procesos": ["Trámite de Grado"]
    },
    "Oficina de Planeación": {
        "script": "Generararchivoexcel_planeacion",
        "procesos": ["Satisfación de estadisticas institucionales"]
    },
    "Dirección de Gestión y Relaciones Internacionales": {
        "script": "Generararchivoexcel_movilidad_entrante",
        "procesos": ["Movilidad Estudiantil Internacional Entrante No 1","Movilidad Estudiantil Internacional Entrante No 2","Gestión de la Movilidad Estudiantil Internacional (saliente)"]
    },
    "Prueba": {
        "script": "Generararchivoexcel_prueba",
        "procesos": ["Prueba grafica"]
    },
    "Almacen": {
        "script": "Generararchivoexcel_Almacen",
        "procesos": ["Entrega de Insumos y Compras Directas - Entrega de Activos"]
    },
    "Sección de Servicios Generales": {
        "script": "Generararchivoexcel_servicio_transporte_seguridad_Aseo",
        "procesos": ["Servicio de transporte","Seguridad en Uninorte","Servicios de Aseo"]
    },
    "Registro": {
        "script": "Generararchivoexcel_certificaciones",
        "procesos": ["Certificaciones académicas por vía Web"]
    },
    "Oficina Genérica / Personalizada": {
        "script": "Generararchivoexcel_generico",
        "procesos": ["Proceso Personalizado"]
    }
}

st.set_page_config(page_title="Exportador de Excel", page_icon="📁", layout="wide")



# --- Estilos personalizados ---
st.markdown("""
    <style>
    .main-title {
        text-align: center;
        font-size: 36px;
        color: #2c3e50;
        margin-top: 20px;
        font-weight: bold;
    }
    .section-title {
        font-size: 20px;
        margin-top: 25px;
        color: #34495e;
        border-bottom: 1px solid #ccc;
        padding-bottom: 6px;
        max-width: 90%;
        margin-left: auto;
        margin-right: auto;
    }
    .footer {
        text-align: center;
        font-size: 13px;
        color: #999;
        margin-top: 30px;
    }
    .block-container {
        max-width: 1200px;
        margin: auto;
    }
    </style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
.main-title {
    font-size: 36px;
    color: #2f3e75;
    text-align: center;
    font-weight: 700;
    margin-top: 30px;
    margin-bottom: 20px;
    background: linear-gradient(to right, #f5f9ff, #e8effc);
    padding: 20px;
    border-radius: 16px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06);
    border: 1px solid #d0d8ec;
    max-width: 900px;
    margin-left: auto;
    margin-right: auto;
}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="main-title">📤 Exportador Automático de Datos a Excel</div>', unsafe_allow_html=True)

# --- Selector de método de procesamiento ---
st.markdown('<div class="section-title">0️⃣ Selecciona el método de procesamiento</div>', unsafe_allow_html=True)
metodo = st.radio("¿Cómo deseas procesar el archivo?", ["Procesar", "Pivotear"], horizontal=True)

# --- Sección: Carga del archivo ---
st.markdown('<div class="section-title">1️⃣ Subir archivo Excel</div>', unsafe_allow_html=True)
archivo_excel = st.file_uploader("Cargar archivo .xlsx", type=["xlsx"])
if archivo_excel:
    st.session_state["archivo_excel"] = archivo_excel
if "archivo_excel" in st.session_state and archivo_excel is None:
    archivo_excel = st.session_state["archivo_excel"]

if archivo_excel is not None:
    try:
        df = pd.read_excel(archivo_excel)
        # Procesar según método seleccionado
        if metodo == "Procesar":
            df = procesar_excel(df)
            st.success("✅ Archivo procesado correctamente (Procesar)")
        elif metodo == "Pivotear":
            df = pivotear_excel(df)
            st.success("✅ Archivo procesado correctamente (Pivotear)")
        st.session_state["df_encuesta"] = df
    except Exception as e:
        st.error(f"❌ Error al leer o procesar el archivo: {e}")
elif "df_encuesta" in st.session_state:
    df = st.session_state["df_encuesta"]

# --- Sección: Configuración ---
st.markdown('<div class="section-title">2️⃣ Seleccionar oficina y parámetros</div>', unsafe_allow_html=True)
oficina_seleccionada = st.selectbox(
    "🏢 Selecciona la oficina",
    options=list(diccionario_oficinas.keys()),
    index=list(diccionario_oficinas.keys()).index(st.session_state.get("oficina_seleccionada", list(diccionario_oficinas.keys())[0]))
)
st.session_state["oficina_seleccionada"] = oficina_seleccionada
procesos_disponibles = diccionario_oficinas[oficina_seleccionada]["procesos"]
valor_guardado = st.session_state.get("proceso_seleccionado", procesos_disponibles[0])
if valor_guardado not in procesos_disponibles:
    valor_guardado = procesos_disponibles[0]

proceso_seleccionado = st.selectbox(
    "🧪 Selecciona el proceso asociado",
    options=procesos_disponibles,
    index=procesos_disponibles.index(valor_guardado)
)
st.session_state["proceso_seleccionado"] = proceso_seleccionado
nombre_archivo = st.text_input("📝 Nombre del archivo de salida (sin extensión)", value="exportado")
numerodepoblacion = st.number_input("👥 Número de población", min_value=1, step=1)

# --- Selección de periodo ---
periodo_unico  = st.text_input("📝 Escribir periodo en que se relizo la encuesta", value="Periodo")

# --- Vista previa del archivo cargado ---
if archivo_excel is not None:
    df = pd.read_excel(archivo_excel)
    st.success("✅ Archivo cargado exitosamente")
    
    # Inicializar variables que se usan en este bloque
    columnas_pregunta_detectadas = []
    columnas_observaciones_detectadas = []
    nombre_columna_general = ""
    #st.dataframe(df.head(), use_container_width=True)

    #Oficina con filtros
    oficina_filtros=["Almacen","Prueba","Sección de Servicios Generales","Registro"]
    proceso_filtros=["Entrega de Insumos y Compras Directas - Entrega de Activos","Prueba grafica","Servicio de transporte","Seguridad en Uninorte","Servicios de Aseo","Certificaciones académicas por vía Web"]
   
    if oficina_seleccionada in oficina_filtros and proceso_seleccionado in proceso_filtros:
        st.markdown('<div class="section-title">🔍 Filtros específicos para prueba</div>', unsafe_allow_html=True)

        if archivo_excel is not None and not df.empty:

            # Filtros eliminados de la barra lateral, solo se muestra la tabla filtrada si aplica
            st.dataframe(df, width='stretch')

        else:
            st.sidebar.warning("⚠️ Sube un archivo Excel para aplicar los filtros.")
    else:   
        st.dataframe(df, width='stretch')
    columnas_todo_no_aplica = [col for col in df.columns if (df[col] == "No Aplica").all()]
    # --- Detectar columnas de preguntas automáticamente ---
    posibles_valores = {"1", "2", "3", "4", "5", "No Aplica"}
    for col in df.columns:
        valores = set(df[col].dropna().astype(str).unique())
        contiene_valores = valores.issubset(posibles_valores) or len(valores.intersection(posibles_valores)) >= 3
        es_general = "general" in col.lower()
        es_numero = "numero" in col.lower()
        if contiene_valores and not es_general and not es_numero and not df[col].isna().all():
            columnas_pregunta_detectadas.append(col)
        elif "general" in col.lower():
            nombre_columna_general = col
    columnas_pregunta_detectadas=[x for x in columnas_pregunta_detectadas if x not in columnas_todo_no_aplica]
    st.markdown('<div class="section-title">🧮 Columnas detectadas como preguntas</div>', unsafe_allow_html=True)
    st.info(f"Preguntas detectadas automáticamente (Sin incluir la pregunta de satisfacción general ): {columnas_pregunta_detectadas}")
    columnas_seleccionadas = st.multiselect("🧾 Selecciona columnas adicionales (opcional)", options=df.columns.tolist(), default=columnas_pregunta_detectadas)

    # --- Detectar columnas de observaciones ---
    palabras_clave_obs = ["comentario", "sugerencia", "observacion"]
    for col in df.columns:
        if any(palabra in col.lower() for palabra in palabras_clave_obs):
            columnas_observaciones_detectadas.append(col)

    st.markdown('<div class="section-title">📝 Columnas de comentarios u observaciones</div>', unsafe_allow_html=True)
    columnas_observaciones = st.multiselect("✏️ Selecciona columnas de observación", options=df.columns.tolist(), default=columnas_observaciones_detectadas)

    #st.markdown(f"**📌 Columna general detectada:** `{nombre_columna_general}`")
    nombre_columna_general=st.multiselect("📌 Selecciona la columna general (opcional)",options=df.columns.tolist(),default=nombre_columna_general)[0]
        
   # Guardar en session_state para que otras páginas puedan acceder
    st.session_state["oficina_seleccionada"]=oficina_seleccionada
    st.session_state["proceso_seleccionado"]=proceso_seleccionado
    st.session_state["df_encuesta"] = df
    st.session_state["columnas_seleccionadas"] = columnas_seleccionadas
    st.session_state["columnas_observaciones"] = columnas_observaciones
    st.session_state["nombre_columna_general"] = nombre_columna_general

    #---Sección de seleccionar columnas para filtros dinámicos (slicers)---------------------------
    st.markdown('<div class="section-title">🎛️ Seleccionar columnas para filtros dinámicos (opcional)</div>', unsafe_allow_html=True)
    st.info("Los filtros dinámicos permiten segmentar los datos en la hoja T+G (ej: por programa, sede, tipo de estudiante, etc.)")
    columnas_filtros_dinamicos = st.multiselect("📊 Selecciona columnas para crear filtros desplegables en T+G:", options=df.columns.tolist())
    st.session_state["columnas_filtros_dinamicos"] = columnas_filtros_dinamicos

    #---Sección de selcionar graficas---------------------------
    st.markdown('<div class="section-title">📊 Seleccionar gráficas (opcional)</div>', unsafe_allow_html=True)
    seleccionadas = st.multiselect("Selecciona las métricas que deseas visualizar:", options=df.columns.tolist())
    # Paso 2: Para cada opción seleccionada, elegir entre 'bar' o 'column'
    tipos_grafica = {}

    for item in seleccionadas:
        tipo = st.selectbox(
            f"Selecciona el tipo de gráfica para '{item}':",
            ['pie', 'column'],
            key=item  # Importante: usar key único para evitar conflictos en el renderizado
        )
        tipos_grafica[item] = tipo
    # --- Sección: Ejecutar proceso ---
st.markdown('<div class="section-title">3️⃣ Ejecutar proceso</div>', unsafe_allow_html=True)
if st.button("🚀 Ejecutar función excel_exportar"):
    if not archivo_excel:
        st.warning("⚠️ Por favor sube el archivo Excel.")
    else:
        try:
            # Obtener datos del session_state
            df = st.session_state.get("df_encuesta", pd.DataFrame())
            preguntas = st.session_state.get("columnas_seleccionadas", [])
            comentarios = st.session_state.get("columnas_observaciones", [])
            general = st.session_state.get("nombre_columna_general", "")
            oficina = oficina_seleccionada
            proceso = proceso_seleccionado
            script_name = diccionario_oficinas[oficina]["script"]
            ruta_script = f"Script de los formatos/{script_name}.py"

            if not os.path.isfile(ruta_script):
                st.error(f"❌ El script '{ruta_script}' no existe.")
            else:
                spec = importlib.util.spec_from_file_location("modulo_dinamico", ruta_script)
                modulo = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(modulo)

                if hasattr(modulo, "excel_exportar"):
                    ruta_salida = f"{nombre_archivo}.xlsx"
                    # Si es Operaciones Tic, poner 'N_Caso' (o la columna que corresponda) de primera
                    if oficina == "Operaciones Tic":
                        col_primera = None
                        for col in df.columns:
                            if col.strip().lower() == "n_caso" or col.strip().lower() == "n caso":
                                col_primera = col
                                break
                        if col_primera:
                            otras = [c for c in df.columns if c != col_primera]
                            df = df[[col_primera] + otras]
                    if script_name == "Generararchivoexcel_generico":
                        filtros_dinamicos = st.session_state.get("columnas_filtros_dinamicos", [])
                        st.info(f"🔍 Filtros dinámicos seleccionados: {len(filtros_dinamicos)} columnas")
                        if filtros_dinamicos:
                            st.write(f"� Columnas: {', '.join(filtros_dinamicos)}")
                        else:
                            st.warning("⚠️ ADVERTENCIA: No hay columnas de filtro seleccionadas. Los slicers NO se crearán.")
                        
                        # Debug adicional
                        import sys
                        print(f"\n{'='*70}")
                        print(f"DEBUG CARGUE.PY - Antes de llamar excel_exportar")
                        print(f"filtros_dinamicos = {filtros_dinamicos}")
                        print(f"Tipo: {type(filtros_dinamicos)}, Longitud: {len(filtros_dinamicos)}")
                        print(f"{'='*70}\n")
                        sys.stdout.flush()
                        
                        with st.spinner("Generando archivo Excel y agregando slicers..."):
                            modulo.excel_exportar(df, nombre_archivo, numerodepoblacion, preguntas, comentarios, general, oficina, proceso, periodo_unico, tipos_grafica, filtros_dinamicos)
                    else:
                        modulo.excel_exportar(df, nombre_archivo, numerodepoblacion, preguntas, comentarios, general, oficina, proceso, periodo_unico, tipos_grafica)
                    st.session_state["ruta_archivo_generado"] = ruta_salida
                    st.success(f"✅ Función ejecutada y archivo generado como '{ruta_salida}'")
                else:
                    st.error("❌ El script no contiene una función llamada 'excel_exportar'.")
        except Exception as e:
            st.error(f"🚨 Error al ejecutar: {e}")

# --- Botón de descarga si se ha generado el archivo ---
if "ruta_archivo_generado" in st.session_state:
    ruta_archivo_generado = st.session_state["ruta_archivo_generado"]
    if os.path.exists(ruta_archivo_generado):
        with open(ruta_archivo_generado, "rb") as f:
            st.download_button(
                label="📥 Descargar archivo generado",
                data=f,
                file_name=os.path.basename(ruta_archivo_generado),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- Pie de página ---
st.markdown('<div class="footer">Desarrollado con ❤️ usando Streamlit · Universidad del Norte · 2025</div>', unsafe_allow_html=True)
