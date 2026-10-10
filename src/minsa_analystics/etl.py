"""
etl.py — Pipeline de consolidación de las bases de dotación de RRHH de MINSA (2019-2025)

Migrado y validado desde notebooks/01_comparacion_columnas.ipynb, donde se
determinó y probó el mapeo canónico de columnas (finales_2019 == ... ==
finales_2025 -> True, 31 columnas).
"""
import pandas as pd
from sqlalchemy import create_engine

# =============================================================================
# Valores para el campo 'es_especialista'
VALORES_SI = {'SI', 'SÍ', '1'}
VALORES_NO = {'NO', '0', '-', ''}   # los nulos también cuentan como NO

# =============================================================================
# Valores para quitar tildes
TILDES = str.maketrans('ÁÉÍÓÚÜáéíóúü', 'AEIOUUaeiouu')   # no incluye la Ñ
COLUMNAS_SIN_TILDES = [
    'provincia', 'descripcionestablecimiento', 'especialidad',
    'regimen_laboral', 'condicion_laboral',
]

# =============================================================================
# Correcciones de nombre de distrito detectadas en la revisión: (ubigeo, nombre erróneo) -> nombre correcto
CORRECCIONES_DISTRITO = {
    ('130112', 'VIR'): 'ALTO TRUJILLO',   # 2024: 7 filas de establecimientos de Alto Trujillo (renaes 00005220 y 00012229)
}

# =============================================================================
CORRECCIONES_QUINTIL = {
    # (ubigeo, año): quintil correcto
    ('130112', 2024): 3,   # 7 filas (renaes 00005220 y 00012229) traían 2; en 2025 es 3 y en El Porvenir era 3
}

# =============================================================================
# Convertimos datos de algunas columnas a Mayusculas
COLUMNAS_MAYUSCULA = ['diresa', 'red', 'microrred', 'categoria']
SIN_RED = {
    'red': 'NO PERTENECE A NINGUNA RED',
    'microrred': 'NO PERTENECE A NINGUNA MICRORED',
}

# =============================================================================
COLUMNAS_A_ELIMINAR = {
    'pea': 'Constante = 1 (contador de tablas dinámicas)',
    'clasificacion': 'Sin actualizar y no usada en reportes; mezcla tipo de establecimiento con unidades administrativas',
}

# =============================================================================
COLUMNAS_BOOLEANAS = ['distfrontera', 'zaf2014final', 'estrategicos']
# =============================================================================
# A. CONFIGURACIÓN ESTRUCTURAL
# =============================================================================

# A.1. Rutas y fila de header por año (relativas a data/raw/)
ARCHIVOS = {
    2019: ("BASE_DIC_2019.xlsx", 5),
    2020: ("BASE_DIC_2020.xlsx", 0),
    2021: ("BASE_DIC_2021.xlsx", 0),
    2022: ("BASE_DIC_2022.xlsx", 0),
    2023: ("BASE_DIC_2023.xlsx", 0),
    2024: ("BASE_DIC_2024.xlsx", 1),
    2025: ("BASE_DIC_2025.xlsx", 0),
}

# A.2. Mapeo canónico de columnas: año -> {nombre_original: nombre_canónico}
# Solo se listan las columnas que necesitan renombrarse; el resto se mantiene igual.
MAPEO_COLUMNAS = {
    2019: {
        'CALSIFICACION': 'CLASIFICACION',
        'DESCRIPCION ESTABLECIMIENTO': 'DESCRIPCIONESTABLECIMIENTO',
        'Dist Frontera': 'DistFrontera',
        'Grupo Final': 'GrupoFinal2',
        'Grupo Final 2': 'GrupoFinal3',
        'PLIEGO + DESCRIP': 'PLIEGODESCRIP',
        'REANES FINAL': 'RENAES',
        'UE + DESCRIP UE': 'UEDESCRIPUE',
        'ZAF 2014 FINAL': 'ZAF2014FINAL',
        'cargo': 'CARGO',
        'id_cargo': 'ID_CARGO',
    },
    2020: {
        'cargo': 'CARGO',
        'Dist Frontera': 'DistFrontera',
        'Grupo Final 2': 'GrupoFinal2',
        'Grupo Final 3': 'GrupoFinal3',
        'id_cargo': 'ID_CARGO',
        'PLIEGO + DESCRIP': 'PLIEGODESCRIP',
        'codigo_renaes': 'RENAES',
        'UE + DESCRIP UE': 'UEDESCRIPUE',
        'ZAF 2014 FINAL': 'ZAF2014FINAL',
        'condicion_laboral rep 2020': 'condicion_laboral',
    },
    2021: {
        'cargo': 'CARGO',
        'edadfinal': 'EDAD',
        'GrupoFinal 2': 'GrupoFinal2',
        'GrupoFinal 3': 'GrupoFinal3',
        'id_cargo': 'ID_CARGO',
        'codigo_renaes': 'RENAES',
        'condicion_laboralrep2020': 'condicion_laboral',
    },
    2022: {
        'cargo': 'CARGO',
        'id_cargo': 'ID_CARGO',
        'codigo_renaes': 'RENAES',
    },
    2023: {
        'cargo': 'CARGO',
        'categoria_establecimiento': 'CATEGORIA',
        'clasificacion': 'CLASIFICACION',
        'departamento': 'DEPARTAMENTO',
        'establecimiento': 'DESCRIPCIONESTABLECIMIENTO',
        'diresa': 'DIRESA',
        'distrito': 'DISTRITO',
        'Dist Frontera': 'DistFrontera',
        'Grupo Final 2': 'GrupoFinal2',
        'Grupo Final 3': 'GrupoFinal3',
        'id_cargo': 'ID_CARGO',
        'microred': 'MICRORRED',
        'PLIEGO Y DESCRIPCION': 'PLIEGODESCRIP',
        'provincia': 'PROVINCIA',
        'red': 'RED',
        'codigo_renaes': 'RENAES',
        'id_ubigeo': 'UBIGEO',
        'UE Y DESCRIPCION': 'UEDESCRIPUE',
        'ZAF 2014 FINAL': 'ZAF2014FINAL',
    },
    2024: {
        'cargo': 'CARGO',
        'DESCRIPCION ESTABLECIMIENTO': 'DESCRIPCIONESTABLECIMIENTO',
        'Dist Frontera': 'DistFrontera',
        'grupo_final_2': 'GrupoFinal2',
        'grupo_final_3': 'GrupoFinal3',
        'id_cargo': 'ID_CARGO',
        'PLIEGO + DESCRIP': 'PLIEGODESCRIP',
        'codigo_renaes': 'RENAES',
        'UE + DESCRIP UE': 'UEDESCRIPUE',
        'ZAF 2014 FINAL': 'ZAF2014FINAL',
    },
    2025: {
        'cargo_recod': 'CARGO',
        'id_cargo_recod': 'ID_CARGO',
        'Edad': 'EDAD',
    },
}

# A.3. Columnas descartadas por año, con el motivo documentado.
# Nota: para cada año, además de esto, se agregan dinámicamente (ver
# inicializar_columnas_descartadas) las columnas EMERGENCIA*, cuyo nombre
# incluye el decreto y cambia cada año.
COLUMNAS_DESCARTADAS = {
    2019: {
        'APS 2015': 'Solo existe en 2019, sin equivalente en años posteriores',
        'CARGO_ESTRUCTURAL': 'No existe en BD 2025 aunque es importante',
        'CODCARGO': 'No es necesario, tiene codigos errados',
        'DESCRIPCION PLIEGO': 'No necesario porque es lo mismo que PLIEGO + DESCRIP',
        'ESTADO': 'No necesario, sin equivalente en años posteriores',
        'INSTITUCION': 'No necesario, sin equivalente en años posteriores',
        'MICRORRED PRIORIZADA APS': 'No necesario, sin equivalente en años posteriores',
        'VRAEM 2016 (DS 040-2016-PCM)': 'No necesario, sin equivalente en años posteriores',
        'VRAEM 2017 (DS 112-2017-PCM)': 'No necesario, sin equivalente en años posteriores',
        'id_condicion_especialidad': 'No necesario, sin equivalente en años posteriores',
        'fecha_nacimiento': 'Transformado a fecha_nacimiento_dt y luego a EDAD; no se conserva',
        'fecha_nacimiento_dt': 'Columna intermedia, usada solo para calcular EDAD',
        'UNIDAD EJECUTORA': 'Duplicado de UE + DESCRIP UE',
        'UE': 'No necesario, UE + DESCRIP UE ya contiene ese dato',
        'PLIEGO': 'No necesario, PLIEGO + DESCRIP ya contiene ese dato',
        'TIPO': 'No necesario, sin equivalente en años posteriores',
    },
    2020: {
        'GRUPO ETAREO': 'No tiene equivalente; se puede obtener de EDAD',
        'Grupo Final 1': 'No existe en la lista canónica',
        'NIVEL': 'No existe en la lista canónica',
        'condicion_laboral formal': 'Igual y hasta más detallado en condicion_laboral',
        'edad final': 'Se calcula EDAD desde fecha_nacimiento; este campo es duplicado y venía sucio',
        'fecha_nacimiento': 'Transformado a fecha_nacimiento_dt y luego a EDAD; no se conserva',
        'fecha_nacimiento_dt': 'Columna intermedia, usada solo para calcular EDAD',
        'TIPO': 'No necesario, sin equivalente en años posteriores',
    },
    2021: {
        'GrupoFinal 1': 'No existe en la lista canónica',
        'NIVEL': 'No existe en la lista canónica',
        'condicion_laboralformal': 'Igual y hasta más detallado en condicion_laboralrep2020',
        'id_condicion_especialidad': 'No existe en la lista canónica',
        'TIPO': 'No necesario, sin equivalente en años posteriores',
    },
    2022: {
        'CATEGORIA 2': 'Duplicado de CATEGORIA',
        'DESCRIPCIONPLIEGO': 'PLIEGODESCRIP ya contiene la misma información',
        'GrupoFinal1': 'No existe en la lista canónica',
        'NIVEL': 'No existe en la lista canónica',
        'PLIEGO': 'Ya contiene la misma información en PLIEGODESCRIP',
        'UE': 'Ya contiene la misma información en UEDESCRIPUE',
        'UNIDADEJECUTORA': 'Ya contiene la misma información en UEDESCRIPUE',
        'fecha_nacimiento': 'Transformado a fecha_nacimiento_dt y luego a EDAD; no se conserva',
        'fecha_nacimiento_dt': 'Columna intermedia, usada solo para calcular EDAD',
        'TIPO': 'No necesario, sin equivalente en años posteriores',
    },
    2023: {
        'Grupo Final 1': 'No existe en la lista canónica',
        'edad': 'Se recalcula con la función calcular_edad para mantener criterio consistente',
        'fecha_nacimiento': 'Transformado a fecha_nacimiento_dt y luego a EDAD; no se conserva',
        'fecha_nacimiento_dt': 'Columna intermedia, usada solo para calcular EDAD',
        'id_condicion_especialidad': 'No existe en la lista canónica',
    },
    2024: {
        'COMUNIDAD_INDIGENA REFERENCIA DGAIN POR UBIGEO FEBRERO 2022': 'No existe en la lista canónica',
        'DESCRIPCION PLIEGO': 'Ya contenido en PLIEGODESCRIP',
        'EESS CLAS JULIO 2022': 'No existe en la lista canónica',
        'Edad': 'Se recalcula con la función calcular_edad para mantener criterio consistente',
        'FRIAJE POR UBIGEO 2022-2024': 'No existe en la lista canónica',
        'Grupo etareo': 'No existe en la lista canónica',
        'HELADAS POR UBIGEO 2022-2024': 'No existe en la lista canónica',
        'Nivel': 'No existe en la lista canónica',
        'PLIEGO': 'Ya contenido en PLIEGODESCRIP',
        'RIS AL 11 NOVIEMBRE 2024': 'No existe en la lista canónica',
        'TIPO': 'No existe en la lista canónica',
        'UE': 'Ya contenido en UEDESCRIPUE',
        'UNIDAD EJECUTORA': 'Ya contenido en UEDESCRIPUE',
        'VRAEM 2022 (DS 133-2022-PCM)': 'No existe en la lista canónica',
        'fecha_nacimiento': 'Transformado a fecha_nacimiento_dt y luego a EDAD; no se conserva',
        'fecha_nacimiento_dt': 'Columna intermedia, usada solo para calcular EDAD',
        'grupo_final_1': 'No existe en la lista canónica',
        'id_condicion_laboral': 'No existe en la lista canónica',
        'id_regimen_laboral': 'No existe en la lista canónica',
        'id_sexo': 'No existe en la lista canónica',
        'nivel_reminerativo_airhsp': 'No existe en la lista canónica',
    },
    2025: {
        'COMUNIDAD_INDIGENAREFERENCIADGAINPORUBIGEOFEBRERO2022': 'No necesario, sin equivalente en años anteriores',
        'DESCRIPCIONPLIEGO': 'Obtenido de PLIEGODESCRIP',
        'DOBLEEMPLEO': 'No necesario, sin equivalente en años anteriores',
        'EESSCLASJULIO2022': 'No necesario, sin equivalente en años anteriores',
        'FRIAJEPORUBIGEO20222024': 'No necesario, sin equivalente en años anteriores',
        'GrupoFinal1': 'Se puede obtener de GrupoFinal2',
        'Grupoetareo': 'No tiene equivalente; se puede obtener de EDAD',
        'HELADASPORUBIGEO20222024': 'No necesario, sin equivalente en años anteriores',
        'NIVEL': 'Se puede obtener de CATEGORIA',
        'RISAL11NOVIEMBRE2024': 'No necesario, sin equivalente en años anteriores',
        'VRAEM2022DS1332022PCM': 'No necesario, sin equivalente en años anteriores',
        'profesion': 'No necesario, sin equivalente en años anteriores',
        'UNIDADEJECUTORA': 'Duplicado de UEDESCRIPUE',
        'UE': 'No necesario, UEDESCRIPUE ya contiene ese dato',
        'PLIEGO': 'No necesario, PLIEGODESCRIP ya contiene ese dato',
        'TIPO': 'No necesario, sin equivalente en años posteriores',
    },
}


# =============================================================================
# B. FUNCIONES DE CARGA Y LIMPIEZA
# =============================================================================

def obtener_columnas_iniciales(ruta, header_row):
    """Lee solo la cabecera del Excel (sin filas) para detectar columnas rápido."""
    df_temp = pd.read_excel(ruta, header=header_row, nrows=0)
    return [str(c).strip() for c in df_temp.columns if c and not str(c).startswith('Unnamed:')]


def cargar_dataframe_limpio(anio, ruta_datos):
    """Carga el Excel completo de un año, forzando IDs como texto y limpiando columnas."""
    nombre_archivo, header_row = ARCHIVOS[anio]
    ruta = f"{ruta_datos}/{nombre_archivo}"

    # Crucial para PostgreSQL: evita que RENAES o IDs pierdan ceros a la izquierda
    dtypes_dict = {
        'REANES FINAL': str, 'RENAES': str, 'codigo_renaes': str,
        'id_cargo': str, 'id_cargo_recod': str, 'CODCARGO': str,
        'UBIGEO': str, 'id_ubigeo': str,
    }

    df = pd.read_excel(ruta, header=header_row, dtype=dtypes_dict)
    df.columns = df.columns.str.strip()

    columnas_validas = [c for c in df.columns if c and not str(c).startswith('Unnamed:')]
    return df[columnas_validas]


def convertir_fecha_nacimiento(df, columna_original='fecha_nacimiento',
                                 formatos=('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y')):
    """
    Convierte fecha_nacimiento a datetime, manejando formatos mixtos entre años,
    texto tipo 'No especifica' (tratado como nulo) y fechas seriales de Excel
    (encontradas en 2022, ej. '34385').
    """
    nueva_columna = f'{columna_original}_dt'

    serie = df[columna_original].replace(
        to_replace=r'(?i)no especifica|n/a|sin dato|no aplica', value=pd.NA, regex=True
    )

    df[nueva_columna] = pd.NaT

    # Fechas seriales de Excel (números puros)
    es_serial = serie.astype(str).str.match(r'^\d+(\.\d+)?$', na=False)
    if es_serial.any():
        df.loc[es_serial, nueva_columna] = pd.to_datetime(
            serie[es_serial].astype(float), unit='D', origin='1899-12-30', errors='coerce'
        )

    # Formatos de texto conocidos, incluyendo variante con hora
    formatos_completos = list(formatos) + ['%Y-%m-%d %H:%M:%S']
    for formato in formatos_completos:
        pendientes = serie.notna() & df[nueva_columna].isna() & ~es_serial
        df.loc[pendientes, nueva_columna] = pd.to_datetime(
            serie.loc[pendientes], format=formato, errors='coerce'
        )

    return df


def calcular_edad(fecha_nacimiento, anio_corte):
    """Calcula la edad entera en años al 31 de diciembre del año de corte."""
    fecha_corte = pd.Timestamp(year=anio_corte, month=12, day=31)
    return (fecha_corte - fecha_nacimiento).dt.days // 365


def descartar_por_prefijo(anio, prefijo, motivo, columnas_del_anio, columnas_descartadas):
    """
    Busca en columnas_del_anio todas las columnas que empiecen con `prefijo`
    (útil para campos cuyo nombre cambia cada año, ej. EMERGENCIA) y las agrega
    a columnas_descartadas[anio] con el motivo indicado. Modifica el diccionario
    in place y devuelve la lista de columnas encontradas.
    """
    encontradas = [c for c in columnas_del_anio if str(c).startswith(prefijo)]
    if anio not in columnas_descartadas:
        columnas_descartadas[anio] = {}
    for col in encontradas:
        columnas_descartadas[anio][col] = motivo
    return encontradas


def inicializar_columnas_descartadas(ruta_datos, mapeo_columnas=None, columnas_descartadas=None):
    """
    Devuelve una copia de COLUMNAS_DESCARTADAS ya completada con las columnas
    EMERGENCIA* de cada año (nombre dinámico por decreto/fecha), detectadas
    leyendo la cabecera real de cada Excel.
    """
    columnas_descartadas = {
        anio: dict(cols) for anio, cols in (columnas_descartadas or COLUMNAS_DESCARTADAS).items()
    }
    for anio, (nombre_archivo, header_row) in ARCHIVOS.items():
        ruta = f"{ruta_datos}/{nombre_archivo}"
        columnas_anio = obtener_columnas_iniciales(ruta, header_row)
        descartar_por_prefijo(
            anio, 'EMERGENCIA',
            'Campo no necesario porque cada 3 meses se actualiza y no está actualizado',
            columnas_anio, columnas_descartadas,
        )
    return columnas_descartadas


def obtener_columnas_finales(columnas_del_anio, anio, mapeo_columnas, columnas_descartadas):
    """Devuelve los nombres definitivos tras aplicar descartes y renombres (para validación)."""
    descartadas = set(columnas_descartadas.get(anio, {}).keys())
    mapeo = mapeo_columnas.get(anio, {})
    mantenidas = set(columnas_del_anio) - descartadas
    return sorted({mapeo.get(col, col) for col in mantenidas})

def consolidar_anio(df, anio, mapeo_columnas, columnas_descartadas):
    """Aplica descartes, renombra columnas al estándar canónico y etiqueta el año."""
    df_proc = df.copy()

    descartar = [c for c in columnas_descartadas.get(anio, {}).keys() if c in df_proc.columns]
    df_proc = df_proc.drop(columns=descartar)

    mapeo = mapeo_columnas.get(anio, {})
    df_proc = df_proc.rename(columns=mapeo)

    # Nombres a minúsculas y snake_case, compatible con PostgreSQL
    df_proc.columns = (
        df_proc.columns.str.lower().str.replace(' ', '_').str.replace('+', 'mas', regex=False)
    )

    df_proc['anio_registro'] = anio
    return df_proc

def normalizar_tipos(df):
    """Unifica tipos mixtos y normaliza es_especialista a 'SI' / 'NO'.
    Detiene el proceso si aparece un valor no reconocido."""
    es = (
        df['es_especialista'].astype('string').str.strip().str.upper()
        .str.replace(r'\.0$', '', regex=True)   # 1.0 -> 1
    )

    # Detener ante valores desconocidos (los nulos son válidos y cuentan como NO)
    mask = es.notna() & ~es.isin(VALORES_SI | VALORES_NO)
    if mask.any():
        detalle = (df.loc[mask, 'anio_registro'].astype(str) + ': ' + es[mask]).value_counts()
        raise ValueError(f"es_especialista: valores no reconocidos (año: valor)\n{detalle}")

    df['es_especialista'] = es.isin(VALORES_SI).map({True: 'SI', False: 'NO'}).astype('string')

    for col in ['id_especialidad', 'especialidad', 'condicion_especialidad']:
        df[col] = (
            df[col].astype('string').str.strip()
            .str.replace(r'\.0$', '', regex=True)
        )
    return df

def limpiar_texto(df):
    """Recorta espacios y colapsa espacios dobles en todas las columnas de texto;
        quita tildes de vocales solo en las columnas donde generaban duplicados (conserva la Ñ).
        Convierte a Mayuscula todos los datos de algunas columnas."""
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].str.strip().str.replace(r'\s+', ' ', regex=True)
    for col in COLUMNAS_SIN_TILDES:
        df[col] = df[col].str.translate(TILDES)
    for col in COLUMNAS_MAYUSCULA:
        df[col] = df[col].str.upper()
    return df

def unificar_sin_red(df):
    """En red y microrred, los nulos, '-' y cualquier variante de 'NO PERTENECE...'
    o '...A NINGUNA...' (incluido el tipeo 'PERNTENE') pasan a la etiqueta única.
    Debe ejecutarse después de limpiar_texto, que deja estas columnas en mayúscula."""
    for col, etiqueta in SIN_RED.items():
        s = df[col]
        sin_red = (
            s.isna()
            | s.isin(['', '-'])
            | s.str.contains('NO PERTENECE|A NINGUNA', na=False)
        )
        df[col] = s.mask(sin_red, etiqueta)
    return df

def corregir_distritos(df):
    """Corrige nombres de distrito erróneos identificados en la validación."""
    for (ubigeo, malo), bueno in CORRECCIONES_DISTRITO.items():
        mask = (df['ubigeo'] == ubigeo) & (df['distrito'] == malo)
        df.loc[mask, 'distrito'] = bueno
    return df

def corregir_quintil(df):
    """El quintil es un atributo del distrito y no cambia entre años; corrige los
    casos donde un ubigeo trae un valor distinto en un año puntual."""
    for (ubigeo, anio), bueno in CORRECCIONES_QUINTIL.items():
        mask = (df['ubigeo'] == ubigeo) & (df['anio_registro'] == anio)
        df.loc[mask, 'quintil'] = bueno
    return df

def unificar_unidades_ejecutoras(df):
    """Cada código de unidad ejecutora usa el nombre del año más reciente en que aparece."""
    cod = df['uedescripue'].str.extract(r'^(\d+)')[0]
    tmp = pd.DataFrame({'cod': cod, 'nombre': df['uedescripue'], 'anio': df['anio_registro']})
    ultimo = tmp.sort_values('anio').drop_duplicates('cod', keep='last').set_index('cod')['nombre']
    df['uedescripue'] = cod.map(ultimo).fillna(df['uedescripue'])
    return df

def unificar_establecimientos(df):
    """Cada renaes usa el nombre del año más reciente en que aparece. Y si tiene comillas quita el 
    ultimo espacio antes de la comilla si es que hubiera. """
    d = 'descripcionestablecimiento'
    df[d] = (
        df[d].str.replace('"', '', regex=False)
        .str.replace(r'\s+', ' ', regex=True)
        .str.strip()
    )
    ultimo = (
        df[['renaes', d, 'anio_registro']]
        .sort_values('anio_registro', kind='stable')
        .drop_duplicates('renaes', keep='last')
        .set_index('renaes')[d]
    )
    df[d] = df['renaes'].map(ultimo)
    return df

def agregar_categoria_actual(df):
    """Conserva categoria (la de cada año) y agrega categoria_actual (la del año
    más reciente por renaes) y flag_cambio_categoria (renaes con >1 categoría)."""
    actual = (
        df[['renaes', 'categoria', 'anio_registro']]
        .sort_values('anio_registro', kind='stable')
        .drop_duplicates('renaes', keep='last')
        .set_index('renaes')['categoria']
    )
    df['categoria_actual'] = df['renaes'].map(actual)
    df['flag_cambio_categoria'] = (
        df.groupby('renaes')['categoria'].transform('nunique') > 1
    )
    return df

def aplicar_reglas_negocio(df):
    """Marca registros que violan las reglas de negocio de es_especialista.
    No modifica los datos originales: solo agrega columnas flag_*."""
    campos_incompletos = (
        df['id_especialidad'].isna()
        | df['especialidad'].isna()
        | df['condicion_especialidad'].isna()
    )
    es_si = df['es_especialista'] == 'SI'
    es_residente = df['condicion_laboral'].astype('string').str.strip().str.lower() == 'residente'

    df['flag_especialista_incompleto'] = es_si & campos_incompletos
    df['flag_residente_especialista'] = es_si & es_residente
    return df

def convertir_a_booleano(df):
    """Convierte a booleano las columnas 0/1. Se detiene si hay nulos o valores
    distintos de 0 y 1, en lugar de convertirlos en silencio."""
    for col in COLUMNAS_BOOLEANAS:
        if df[col].isna().any():
            raise ValueError(f'{col}: tiene nulos; no se puede convertir a booleano')
        inesperados = set(df[col].unique()) - {0, 1}
        if inesperados:
            raise ValueError(f'{col}: valores inesperados {inesperados}')
        df[col] = df[col].astype(bool)
    return df

def eliminar_columnas_innecesarias(df):
    """Elimina las columnas descartadas durante la validación del consolidado."""
    return df.drop(columns=list(COLUMNAS_A_ELIMINAR))

# =============================================================================
# C. PIPELINE PRINCIPAL
# =============================================================================

def consolidar_datos(ruta_datos="data/raw"):
    """
    Carga, limpia y consolida los 7 años (2019-2025) en un único DataFrame,
    aplicando el mapeo canónico de columnas ya validado.
    """
    columnas_descartadas = inicializar_columnas_descartadas(ruta_datos)

    dataframes = []
    for anio in ARCHIVOS:
        df = cargar_dataframe_limpio(anio, ruta_datos)

        # 2021 y 2025 ya traen la edad calculada por MINSA (edadfinal / Edad),
        # sin columna de fecha de nacimiento; los demás años sí la traen cruda.
        if 'fecha_nacimiento' in df.columns:
            df = convertir_fecha_nacimiento(df)
            df['EDAD'] = calcular_edad(df['fecha_nacimiento_dt'], anio)

        df_final = consolidar_anio(df, anio, MAPEO_COLUMNAS, columnas_descartadas)
        dataframes.append(df_final)
        print(f"{anio}: {len(df_final):,} filas, {len(df_final.columns)} columnas")

    consolidado = pd.concat(dataframes, ignore_index=True)
    consolidado = normalizar_tipos(consolidado)
    consolidado = limpiar_texto(consolidado)
    consolidado = unificar_sin_red(consolidado)
    consolidado = corregir_distritos(consolidado)
    consolidado = corregir_quintil(consolidado)
    consolidado = unificar_unidades_ejecutoras(consolidado)
    consolidado = unificar_establecimientos(consolidado)
    consolidado = agregar_categoria_actual(consolidado)   
    consolidado = aplicar_reglas_negocio(consolidado)  
    consolidado = convertir_a_booleano(consolidado) 
    consolidado = eliminar_columnas_innecesarias(consolidado)  


    
    print(f"\nConsolidado final: {len(consolidado):,} filas, {len(consolidado.columns)} columnas")
    return consolidado


def exportar_a_postgresql(df_consolidado, tabla_destino, usuario, password, host, puerto, bd):
    """Carga masiva del dataset unificado a PostgreSQL."""
    str_conexion = f'postgresql://{usuario}:{password}@{host}:{puerto}/{bd}'
    engine = create_engine(str_conexion)

    print(f"Iniciando carga en la tabla '{tabla_destino}'...")
    df_consolidado.to_sql(name=tabla_destino, con=engine, if_exists='replace', index=False, chunksize=5000)
    print("Carga masiva finalizada con éxito en PostgreSQL.")


if __name__ == "__main__":
    df = consolidar_datos()