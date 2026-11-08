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
TILDES = str.maketrans('ÁÉÍÓÚÜáéíóúüÀÈÌÒÙàèìòù', 'AEIOUUaeiouuAEIOUaeiou')   # no incluye la Ñ


COLUMNAS_SIN_TILDES = [
    'provincia', 'descripcionestablecimiento', 'especialidad',
    'regimen_laboral', 'condicion_laboral', 'grupofinal2', 'grupofinal3'
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
COLUMNAS_MAYUSCULA = ['diresa', 'red', 'microrred', 'categoria', 'condicion_especialidad', 'especialidad','regimen_laboral']

SIN_RED = {
    'red': 'NO PERTENECE A NINGUNA RED',
    'microrred': 'NO PERTENECE A NINGUNA MICRORED',
}

# =============================================================================
VALORES_A_NULO = {   # columna: valores que equivalen a "sin especialidad"
    'id_especialidad': ['-', '0'],
    'especialidad': ['-', '0'],
    'condicion_especialidad': ['-', '0'],
}

# =============================================================================
COLUMNAS_A_ELIMINAR = {
    'pea': 'Constante = 1 (contador de tablas dinámicas)',
    'clasificacion': 'Sin actualizar y no usada en reportes; mezcla tipo de establecimiento con unidades administrativas',
    'id_cargo': 'El campo no tiene corregido los codigos y hay inconsistencia con el cargo'
}

# =============================================================================
COLUMNAS_BOOLEANAS = ['distfrontera', 'zaf2014final', 'estrategicos']

# =============================================================================
VALORES_SIN_DATO = {
    'sexo': 'Sin dato',   # 543 nulos de 2019 (médicos residentes sin sexo ni edad en la fuente)
}

# =============================================================================
# Cantidad esperada de filas con id_especialidad pero sin nombre (texto truncado en la fuente)
# Si se añade nuevo excel y se ve que hay mas de 12 se detentra y enviara alerta
IDS_SIN_NOMBRE_ESPERADOS = 12

# =============================================================================
# Residentes y especialidad
CONDICIONES_RESIDENTE = {'residente', 'residente plaza cautiva'}   # en minúscula

# nombre de especialidad -> id (solo especialidades que existen en una única profesión)
CODIGOS_ESPECIALIDAD = {
    'FARMACIA HOSPITALARIA': 'QF08',   # 15 filas de residentes QF de 2021 sin id
    'FARMACIA CLINICA': 'QF04',        # 9 filas
}
SIN_DATO_ESPECIALIDAD = 'SIN_DATO'

COND_SIN_ESPECIALIDAD_ESPERADAS = 777   # 771 + los 6 'NO ESPECIFICADO' sin especialidad
ESP_SIN_CONDICION_ESPERADAS = 75        # 68 SI incompletos + 7 NO

CONDICION_RESIDENTE_SIN_SER_RESIDENTE_ESPERADAS = 1   # médico 2020 CAS COVID con condicion_especialidad = RESIDENTE
NO_ESPECIALISTAS_CON_ESPECIALIDAD_ESPERADOS = 104     # 2020-2021; NO con especialidad y condición

# =======================================================================
# Quitar especialidad a grupo no profesionales
GRUPOS_SIN_ESPECIALIDAD = {
    'Tecnico Asistencial', 'Tecnico Administrativo',
    'Auxiliar Asistencial', 'Auxiliar Administrativo',
}
TECNICOS_AUX_CON_ESPECIALIDAD_ESPERADOS = 206   # 2019-2021 y 2023

# =======================================================================
# Regimenes laborales esperados
REGIMENES_ESPERADOS = {
    'REGIMEN 276', 'REGIMEN 1057 (CAS)', 'REGIMEN 728',
    'SERVICIO DE TERCEROS / LOCACION DE SERVICIOS',
    'INTERNOS DE CARRERAS DE LA SALUD', 'PRACTICANTES',
    'PERSONAL PPTO CONTRALORIA', 'PERSONAL PPTO SERVIR',
    'FAC O PAC / PPTO MEF',
}

# =======================================================================
# Condiciones laborales esperados
MAPEO_CONDICION = {
    # (regimen_laboral, condicion_laboral): (condicion_laboral_1 [agrupada], condicion_laboral_2 [detalle])
    ('FAC O PAC / PPTO MEF', 'FAC O PAC / PPTO MEF'): ('FAC O PAC / PPTO MEF', 'FAC O PAC / PPTO MEF'),
    ('INTERNOS DE CARRERAS DE LA SALUD', 'CON ESTIPENDIO (PAGADO)'): ('INTERNOS DE CARRERAS DE LA SALUD', 'CON ESTIPENDIO (PAGADO)'),
    ('INTERNOS DE CARRERAS DE LA SALUD', 'SIN ESTIPENDIO (NO PAGADO)'): ('INTERNOS DE CARRERAS DE LA SALUD', 'SIN ESTIPENDIO (NO PAGADO)'),
    ('PERSONAL PPTO CONTRALORIA', 'PERSONAL PPTO CONTRALORIA'): ('PERSONAL PPTO CONTRALORIA', 'PERSONAL PPTO CONTRALORIA'),
    ('PERSONAL PPTO SERVIR', 'PERSONAL PPTO SERVIR'): ('PERSONAL PPTO SERVIR', 'PERSONAL PPTO SERVIR'),
    ('PRACTICANTES', 'PRACTICANTES'): ('PRACTICANTES', 'PRACTICANTES'),
    # CAS
    ('REGIMEN 1057 (CAS)', 'CONTRATO CAS'): ('CONTRATO CAS', 'CONTRATO CAS'),
    ('REGIMEN 1057 (CAS)', 'CAS REGULAR MINSA'): ('CONTRATO CAS', 'CONTRATO CAS'),
    ('REGIMEN 1057 (CAS)', 'REINCORPORADO'): ('CONTRATO CAS', 'CONTRATO CAS'),
    ('REGIMEN 1057 (CAS)', 'CONTRATO MUNICIPAL CAS'): ('CONTRATO CAS', 'CONTRATO CAS'),
    ('REGIMEN 1057 (CAS)', 'CAS LEY 31538'): ('CONTRATO CAS', 'LEY 31538'),
    ('REGIMEN 1057 (CAS)', 'CAS LEY N° 31538'): ('CONTRATO CAS', 'LEY 31538'),
    ('REGIMEN 1057 (CAS)', 'CONTRATADO A PLAZO DETERMINADO'): ('CONTRATO CAS', 'CONTRATADO A PLAZO DETERMINADO'),
    ('REGIMEN 1057 (CAS)', 'CONTRATADO A PLAZO INDETERMINADO'): ('CONTRATO CAS', 'CONTRATADO A PLAZO INDETERMINADO'),
    ('REGIMEN 1057 (CAS)', 'CAS COVID MINSA'): ('CONTRATO CAS', 'CAS COVID'),
    ('REGIMEN 1057 (CAS)', 'CAS COVID HOSPITALES MODULARES'): ('CONTRATO CAS', 'CAS COVID'),
    ('REGIMEN 1057 (CAS)', 'CAS COVID BRIGADAS ENA LLAMOSAS'): ('CONTRATO CAS', 'CAS COVID'),
    ('REGIMEN 1057 (CAS)', 'CAS COVID BRIGADAS CUBANOS'): ('CONTRATO CAS', 'CAS COVID'),
    ('REGIMEN 1057 (CAS)', 'NO ESPECIFICA'): ('NO ESPECIFICA', 'NO ESPECIFICA'),
    # Regimen 276
    ('REGIMEN 276', 'NOMBRADO'): ('NOMBRADO', 'NOMBRADO'),
    ('REGIMEN 276', 'CONTRATADO 276 - PLAZO FIJO'): ('CONTRATADO 276 - PLAZO FIJO', 'CONTRATADO 276 - PLAZO FIJO'),
    ('REGIMEN 276', 'CONTRATADO 276 - PLAZO INDETERMINADO'): ('CONTRATADO 276 - PLAZO INDETERMINADO', 'CONTRATADO 276 - PLAZO INDETERMINADO'),
    ('REGIMEN 276', 'RESIDENTE'): ('RESIDENTE', 'RESIDENTE'),
    ('REGIMEN 276', 'RESIDENTE PLAZA CAUTIVA'): ('RESIDENTE', 'RESIDENTE PLAZA CAUTIVA'),
    ('REGIMEN 276', 'DESTACADO'): ('DESTACADO', 'DESTACADO'),
    ('REGIMEN 276', 'SERUMS REMUNERADO'): ('SERUMS', 'SERUMS REMUNERADO'),
    ('REGIMEN 276', 'SERUMS EQUIVALENTE'): ('SERUMS', 'SERUMS EQUIVALENTE'),
    ('REGIMEN 276', 'NO ESPECIFICA'): ('NO ESPECIFICA', 'NO ESPECIFICA'),
    ('REGIMEN 276', 'DESIGNADO DE CARRERA'): ('DESIGNADO', 'DESIGNADO DE CARRERA'),
    ('REGIMEN 276', 'DESIGNADO NO DE CARRERA'): ('DESIGNADO', 'DESIGNADO NO DE CARRERA'),
    ('REGIMEN 276', 'DESIGNADO'): ('DESIGNADO', 'DESIGNADO'),
    ('REGIMEN 276', 'REINCORPORADO'): ('REINCORPORADO', 'REINCORPORADO'),
    ('REGIMEN 276', 'ENCARGADO'): ('ENCARGADO', 'ENCARGADO'),
    # Regimen 728
    ('REGIMEN 728', 'CONTRATO CLAS'): ('CONTRATO CLAS', 'CONTRATO CLAS'),
    ('REGIMEN 728', 'CONTRATO 728 - PLAZO INDETERMINADO'): ('CONTRATO 728 - PLAZO INDETERMINADO', 'CONTRATO 728 - PLAZO INDETERMINADO'),
    ('REGIMEN 728', 'DESIGNADO NO DE CARRERA'): ('DESIGNADO', 'DESIGNADO NO DE CARRERA'),
    ('REGIMEN 728', 'DESIGNADO'): ('DESIGNADO', 'DESIGNADO'),
    ('REGIMEN 728', 'NO ESPECIFICA'): ('NO ESPECIFICA', 'NO ESPECIFICA'),
    ('REGIMEN 728', 'REINCORPORADO'): ('REINCORPORADO', 'REINCORPORADO'),
    # Terceros
    ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'SERVICIO DE TERCEROS / LOCACION DE SERVICIOS'):
        ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'SERVICIO DE TERCEROS / LOCACION DE SERVICIOS'),
    ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'CONTRATO MUNICIPAL POR SERVICIO DE TERCEROS / POR LOCACION DE SERVICIOS'):
        ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'CONTRATO MUNICIPAL POR SERVICIO DE TERCEROS / POR LOCACION DE SERVICIOS'),
    ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'FONDO MINERO'):
        ('SERVICIO DE TERCEROS / LOCACION DE SERVICIOS', 'FONDO MINERO'),
}

# =======================================================================
# Corrigiendo descripcion de cargos

CORRECCIONES_CARGO = {
    # Erratas y nombre incompleto
    'PROFESONAL ADMINISTRATIVO NO ESPECIFICADO': 'PROFESIONAL ADMINISTRATIVO NO ESPECIFICADO',
    'AUXIILIAR ASISTENCIAL NO ESPECIFICADO'    : 'AUXILIAR ASISTENCIAL NO ESPECIFICADO',
    'AUXIILIAR ADMINISTRATIVO NO ESPECIFICADO' : 'AUXILIAR ADMINISTRATIVO NO ESPECIFICADO',
    'PROFESIONAL ASISTENCIAL NO ESPECIFICA'    : 'PROFESIONAL DE LA SALUD NO ESPECIFICADO',
    # CNN9: tres nombres para el mismo cargo
    'TECNICO ASISTENCIAL NO ESPECIFICADO'      : 'TECNICO/A ASISTENCIAL',
    'TECNICO ASISTENCIAL'                      : 'TECNICO/A ASISTENCIAL',
    # CA114.x: nombre vigente desde 2023 (sin paréntesis)
    'TECNICO/A SANITARIO/A V (TECNICO/A DE SERVICIOS GENERALES)' : 'TECNICO/A SANITARIO/A V',
    'TECNICO/A SANITARIO/A IV TECNICO/A DE SERVICIOS GENERALES)' : 'TECNICO/A SANITARIO/A IV',
    'TECNICO/A SANITARIO/A III (TECNICO/A DE SERVICIOS GENERALES': 'TECNICO/A SANITARIO/A III',
    'TECNICO/A SANITARIO/A II (TECNICO/A DE SERVICIOS GENERALES)': 'TECNICO/A SANITARIO/A II',
    'TECNICO/A SANITARIO/A I (TECNICO/A DE SERVICIOS GENERALES)' : 'TECNICO/A SANITARIO/A I',
    # CA150.x y CA151.x
    'ASISTENTE EN SERVICIO DE SALUD II (ASISTENTE/A PROFESIONAL II)' : 'ASISTENTE EN SERVICIO DE SALUD II',
    'ASISTENTE EN SERVICIOS SOCIALES II (ASISTENTE/A PROFESIONAL II)': 'ASISTENTE EN SERVICIOS SOCIALES II',
    'ASISTENTE EN RECURSOS NATURALES II (ASISTENTE/A PROFESIONAL II)': 'ASISTENTE EN RECURSOS NATURALES II',
    'ASISTENTE EN SERVICIO DE SALUD I (ASISTENTE/A PROFESIONAL I)'   : 'ASISTENTE EN SERVICIO DE SALUD I',
    'ASISTENTE EN SERVICIOS SOCIALES I (ASISTENTE/A PROFESIONAL I)'  : 'ASISTENTE EN SERVICIOS SOCIALES I',
    'ASISTENTE EN RECURSOS NATURALES I (ASISTENTE/A PROFESIONAL I)'  : 'ASISTENTE EN RECURSOS NATURALES I',
}
FILAS_CARGO_CORREGIDAS_ESPERADAS = 12509

# =======================================================================
# Corrigiendo Grupofinal2 y Grupofinal3

GRUPOFINAL3_POR_GRUPOFINAL2 = {
    'Auxiliar Administrativo'   : {'Auxiliar Administrativo'},
    'Auxiliar Asistencial'      : {'Auxiliar Asistencial'},
    'Profesional Administrativo': {'Profesional Administrativo'},
    'Tecnico Administrativo'    : {'Tecnico Administrativo'},
    'Tecnico Asistencial'       : {'Tecnico Asistencial'},
    'Profesional Asistencial'   : {
        'Biologo', 'Enfermero', 'Ingeniero Sanitario', 'Medico', 'Medico Veterinario',
        'Nutricionista', 'Obstetra', 'Odontologo', 'Profesional Asistencial No Especificado',
        'Psicologo', 'Quimico', 'Quimico Farmaceutico', 'Tecnicos Especializados',
        'Tecnologo Medico', 'Trabajadora Social',
    },
}

# =======================================================================
# Creando SubGrupo(Asistencial, Administrativo)

SUBGRUPO_POR_GRUPO = {
    'Medico'                    : 'Asistencial', 
    'Profesional Asistencial'   : 'Asistencial',
    'Tecnico Asistencial'       : 'Asistencial',
    'Auxiliar Asistencial'      : 'Asistencial',
    'Profesional Administrativo': 'Administrativo', 
    'Tecnico Administrativo'    : 'Administrativo',
    'Auxiliar Administrativo'   : 'Administrativo',
}

# =======================================================================
# Condiciones laboral 1 que no deben tener especialidad
CONDICIONES_EN_FORMACION = {'SERUMS', 'INTERNOS DE CARRERAS DE LA SALUD'}
ESPECIALISTAS_EN_FORMACION_ESPERADOS = 4

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
    """Edad cumplida al 31 de diciembre del año de corte. A esa fecha todos ya cumplieron años,
    así que equivale a anio_corte - año de nacimiento."""
    return anio_corte - fecha_nacimiento.dt.year

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

def unificar_regimen_laboral(df):
    """Unifica la grafía de FAC o PAC ('–' → '/') y se detiene si aparece un régimen
    nuevo, para revisarlo. Va después de limpiar_texto, que deja la columna en mayúscula."""
    df['regimen_laboral'] = df['regimen_laboral'].str.replace('–', '/', regex=False)
    nuevos = set(df['regimen_laboral'].dropna().unique()) - REGIMENES_ESPERADOS
    if nuevos or df['regimen_laboral'].isna().any():
        raise ValueError(f"regimen_laboral: valores no esperados {nuevos}; revisar antes de continuar")
    return df

def corregir_cargo(df):
    """Unifica erratas y nombres renombrados de cargo (queda el nombre vigente). Además corrige
    la única fila con id CA173 cuyo cargo decía ESPECIALISTA ADMINISTRATIVO II, ya que su grupo
    (Químico Farmacéutico) confirma el id. Debe ir antes de eliminar id_cargo."""
    n = int(df['cargo'].isin(CORRECCIONES_CARGO).sum())
    if n != FILAS_CARGO_CORREGIDAS_ESPERADAS:
        raise ValueError(f"corregir_cargo: {n} filas (esperadas {FILAS_CARGO_CORREGIDAS_ESPERADAS}); "
                         "revisar antes de continuar")
    df['cargo'] = df['cargo'].replace(CORRECCIONES_CARGO)

    m = (df['id_cargo'].str.upper() == 'CA173') & (df['cargo'] == 'ESPECIALISTA ADMINISTRATIVO II')
    if int(m.sum()) != 1:
        raise ValueError(f"corregir_cargo: {int(m.sum())} filas CA173 (esperada 1)")
    df.loc[m, 'cargo'] = 'QUIMICO FARMACEUTICO'
    return df

def unificar_grupos(df):
    """Unifica la escritura de grupofinal3 (las tildes las quita limpiar_texto) y valida la
    jerarquía grupofinal2 -> grupofinal3. Se detiene si aparece un grupo o un cruce nuevo."""
    df['grupofinal3'] = df['grupofinal3'].replace({
        'Ingeniero sanitario': 'Ingeniero Sanitario',
        'Profesional Asistencial no especificado': 'Profesional Asistencial No Especificado',
    })
    pares = set(map(tuple, df[['grupofinal2', 'grupofinal3']].drop_duplicates().to_numpy()))
    esperados = {(g2, g3) for g2, g3s in GRUPOFINAL3_POR_GRUPOFINAL2.items() for g3 in g3s}
    if df[['grupofinal2', 'grupofinal3']].isna().any().any() or pares - esperados:
        raise ValueError(f"unificar_grupos: cruces no esperados {pares - esperados}; "
                         "revisar antes de continuar")
    return df

def agregar_grupofinal4_y_subgrupo(df):
    """grupofinal4 = grupofinal2 con los médicos separados (según grupofinal3);
    sub_grupo = Asistencial / Administrativo. Se detiene si hay un grupo sin subgrupo
    o un médico fuera de Profesional Asistencial."""
    es_medico = df['grupofinal3'] == 'Medico'          # igualdad exacta, no contains
   
    if not es_medico.any():
        raise ValueError("grupofinal4: no se encontró ningún 'Médico' en grupofinal3")
    if (es_medico & (df['grupofinal2'] != 'Profesional Asistencial')).any():
        raise ValueError("grupofinal4: hay médicos fuera de Profesional Asistencial")

    df['grupofinal4'] = df['grupofinal2'].mask(es_medico, 'Medico')
    df['sub_grupo'] = df['grupofinal4'].map(SUBGRUPO_POR_GRUPO)
    if df['sub_grupo'].isna().any():
        sin = df.loc[df['sub_grupo'].isna(), 'grupofinal4'].unique()
        raise ValueError(f"sub_grupo: grupos sin mapeo {list(sin)}")
    return df

def agregar_condiciones_laborales(df):
    """Agrega condicion_laboral_1 (agrupada) y condicion_laboral_2 (detalle) según
    MAPEO_CONDICION. Los SERUMS se clasifican por tipo: EQUIVALENTE si lo dice el texto,
    REMUNERADO en los demás casos. Se detiene si aparece una combinación sin mapeo.
    No modifica condicion_laboral: las funciones de residentes la siguen usando."""
    reg = df['regimen_laboral'].str.upper().str.replace('–', '/', regex=False)
    cond = df['condicion_laboral'].str.upper().str.replace('–', '/', regex=False)
    es_serums = cond.str.startswith('SERUMS', na=False)
    tipo = cond.str.contains('EQUIVALENTE', na=False).map({True: 'EQUIVALENTE', False: 'REMUNERADO'})
    cond = cond.mask(es_serums, 'SERUMS ' + tipo)

    clave = reg + '|' + cond
    df['condicion_laboral_1'] = clave.map({f'{r}|{c}': v[0] for (r, c), v in MAPEO_CONDICION.items()})
    df['condicion_laboral_2'] = clave.map({f'{r}|{c}': v[1] for (r, c), v in MAPEO_CONDICION.items()})

    sin_mapeo = clave[df['condicion_laboral_1'].isna()].unique()
    if len(sin_mapeo):
        raise ValueError(f"condicion_laboral: combinaciones sin mapeo {list(sin_mapeo)}; revisar antes de continuar")
    return df

def rellenar_sin_dato(df):
    """Reemplaza los nulos de columnas categóricas por una etiqueta explícita."""
    for col, etiqueta in VALORES_SIN_DATO.items():
        df[col] = df[col].fillna(etiqueta)
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

def limpiar_especialidad(df):
    """Unifica 'sin especialidad' ('-', '0' y nulos) como nulo y las variantes de
    'no especifica' de condicion_especialidad. Va después de limpiar_texto, que deja
    esa columna en mayúscula."""
    for col, valores in VALORES_A_NULO.items():
        s = df[col]
        df[col] = s.mask(s.isin(valores), pd.NA)
    s = df['condicion_especialidad']
    df['condicion_especialidad'] = s.mask(s.str.contains('ESPECIF', na=False), 'NO ESPECIFICADO')
    return df

def corregir_ids_sin_nombre(df):
    """Filas con id_especialidad pero sin nombre de especialidad (texto truncado en la
    fuente, ej. 'GESTIO'): pasan a es_especialista = NO y se vacían id_especialidad
    y condicion_especialidad. Se detiene si la cantidad cambia, para revisarlo."""
    mask = df['id_especialidad'].notna() & df['especialidad'].isna()
    if mask.sum() != IDS_SIN_NOMBRE_ESPERADOS:
        raise ValueError(
            f"id_especialidad sin nombre: {mask.sum()} filas "
            f"(esperadas {IDS_SIN_NOMBRE_ESPERADOS}); revisar antes de continuar"
        )
    df.loc[mask, 'es_especialista'] = 'NO'
    df.loc[mask, ['id_especialidad', 'condicion_especialidad']] = pd.NA
    return df

def unificar_especialidades(df):
    """Cada id_especialidad usa el nombre del año más reciente en que aparece.
    Corrige renombres entre años y errores de digitación (p. ej. N005, TMLC9).
    Va después de limpiar_texto, que deja el nombre en mayúscula y sin tildes."""
    validas = df.dropna(subset=['id_especialidad', 'especialidad'])
    ultimo = (
        validas.sort_values('anio_registro', kind='stable')
        .drop_duplicates('id_especialidad', keep='last')
        .set_index('id_especialidad')['especialidad']
    )
    nuevo = df['id_especialidad'].map(ultimo)
    df['especialidad'] = nuevo.fillna(df['especialidad']).astype('string')
    return df

def _es_residente(df):
    return (df['condicion_laboral'].astype('string').str.strip().str.lower()
            .isin(CONDICIONES_RESIDENTE))

def completar_ids_especialidad(df):
    """Completa id_especialidad con un diccionario explícito nombre -> id."""
    sin_id = df['especialidad'].notna() & df['id_especialidad'].isna()
    ids = df['especialidad'].map(CODIGOS_ESPECIALIDAD)
    mask = sin_id & ids.notna()
    df.loc[mask, 'id_especialidad'] = ids[mask]
    print(f"completar_ids_especialidad: {int(mask.sum())} filas con id completado")
    return df

def corregir_residentes(df):
    """Un residente está en formación: es_especialista = NO y condicion_especialidad = RESIDENTE
    (la fuente de residentes venía de otro Excel y no siempre era coherente)."""
    es_res = _es_residente(df)
    cambia_es = es_res & (df['es_especialista'] != 'NO')
    cambia_cond = es_res & (df['condicion_especialidad'] != 'RESIDENTE').fillna(True)
    df.loc[es_res, 'es_especialista'] = 'NO'
    df.loc[es_res, 'condicion_especialidad'] = 'RESIDENTE'
    print(f"corregir_residentes: es_especialista -> NO en {int(cambia_es.sum())}; "
          f"condicion_especialidad -> RESIDENTE en {int(cambia_cond.sum())}")
    return df

def igualar_condicion_especialidad(df):
    """Hace coincidir condicion_especialidad con especialidad (excepto residentes).

    1. Condición sin especialidad (registro incompleto en el sistema de origen):
       condicion_especialidad pasa a nulo.
    2. Especialidad sin condición: condicion_especialidad = 'NO ESPECIFICADO',
       es_especialista = 'SI' y se marca flag_especialista_incompleto.
    Se detiene si las cantidades cambian, para revisarlo.
    Debe ir después de corregir_residentes y antes de rellenar_sin_dato_residentes.
    """
    es_res = _es_residente(df)

    # 1. Calcular las filas afectadas (todavía sin modificar nada)
    sin_esp = df['condicion_especialidad'].notna() & df['especialidad'].isna() & ~es_res
    sin_cond = df['especialidad'].notna() & df['condicion_especialidad'].isna() & ~es_res

    # 2. Detener si las cantidades no son las revisadas
    n_sin_esp, n_sin_cond = int(sin_esp.sum()), int(sin_cond.sum())
    if n_sin_esp != COND_SIN_ESPECIALIDAD_ESPERADAS or n_sin_cond != ESP_SIN_CONDICION_ESPERADAS:
        raise ValueError(
            f"igualar_condicion_especialidad: condición sin especialidad = {n_sin_esp} "
            f"(esperadas {COND_SIN_ESPECIALIDAD_ESPERADAS}); especialidad sin condición = {n_sin_cond} "
            f"(esperadas {ESP_SIN_CONDICION_ESPERADAS}); revisar antes de continuar"
        )

    # 3. Aplicar las correcciones
    df.loc[sin_esp, 'condicion_especialidad'] = pd.NA

    cambia_si = sin_cond & (df['es_especialista'] != 'SI')
    df.loc[sin_cond, 'condicion_especialidad'] = 'NO ESPECIFICADO'
    df.loc[sin_cond, 'es_especialista'] = 'SI'
    df.loc[sin_cond, 'flag_especialista_incompleto'] = True

    print(f"igualar_condicion_especialidad: {n_sin_esp} condiciones a nulo; "
          f"{n_sin_cond} -> NO ESPECIFICADO ({int(cambia_si.sum())} pasan a es_especialista = SI)")
    return df

def corregir_especialistas_no_marcados(df):
    """1. condicion_especialidad = RESIDENTE solo corresponde a residentes por condicion_laboral;
       en quien no lo es, pasa a NO ESPECIFICADO.
    2. Quien no es residente y tiene especialidad con condición (RNE, TITULO, CONSTANCIA o
       NO ESPECIFICADO) es especialista: es_especialista pasa de NO a SI.
    Se detiene si las cantidades cambian, para revisarlo."""
    es_res = _es_residente(df)

    # Calcular las filas afectadas (todavía sin modificar nada)
    falso_res = (df['condicion_especialidad'] == 'RESIDENTE') & ~es_res
    mask = (
        (df['es_especialista'] == 'NO')
        & df['especialidad'].notna()
        & df['condicion_especialidad'].notna()
        & ~es_res
    )

    n_falso, n = int(falso_res.sum()), int(mask.sum())
    if (n_falso != CONDICION_RESIDENTE_SIN_SER_RESIDENTE_ESPERADAS
            or n != NO_ESPECIALISTAS_CON_ESPECIALIDAD_ESPERADOS):
        raise ValueError(
            f"corregir_especialistas_no_marcados: RESIDENTE sin ser residente = {n_falso} "
            f"(esperadas {CONDICION_RESIDENTE_SIN_SER_RESIDENTE_ESPERADAS}); "
            f"NO con especialidad = {n} (esperadas {NO_ESPECIALISTAS_CON_ESPECIALIDAD_ESPERADOS}); "
            "revisar antes de continuar"
        )

    # Aplicar las correcciones
    df.loc[falso_res, 'condicion_especialidad'] = 'NO ESPECIFICADO'
    df.loc[mask, 'es_especialista'] = 'SI'
    print(f"corregir_especialistas_no_marcados: {n_falso} RESIDENTE -> NO ESPECIFICADO; "
          f"{n} filas pasan a es_especialista = SI")
    return df

def quitar_especialidad_tecnicos_auxiliares(df):
    """Un técnico o auxiliar no ejerce la especialidad por el cargo que ocupa: pasa a
    es_especialista = NO y se vacían los campos de especialidad. Se deja flag para
    conservar trazabilidad. Se detiene si la cantidad cambia, para revisarlo.
    Debe ir después de corregir_especialistas_no_marcados."""
    mask = df['grupofinal2'].isin(GRUPOS_SIN_ESPECIALIDAD) & (df['es_especialista'] == 'SI')
    n = int(mask.sum())
    if n != TECNICOS_AUX_CON_ESPECIALIDAD_ESPERADOS:
        raise ValueError(
            f"quitar_especialidad_tecnicos_auxiliares: {n} filas "
            f"(esperadas {TECNICOS_AUX_CON_ESPECIALIDAD_ESPERADOS}); revisar antes de continuar"
        )
    df['flag_especialidad_removida_por_cargo'] = mask
    df.loc[mask, 'es_especialista'] = 'NO'
    df.loc[mask, ['id_especialidad', 'especialidad', 'condicion_especialidad']] = pd.NA
    df.loc[mask, 'flag_especialista_incompleto'] = False
    print(f"quitar_especialidad_tecnicos_auxiliares: {n} filas pasan a es_especialista = NO")
    return df

def quitar_especialidad_en_formacion(df):
    """SERUMS e internos están en formación, así que no pueden ser especialistas: pasan a
    es_especialista = NO y se vacían los campos de especialidad. Conserva trazabilidad con un flag
    y se detiene si la cantidad cambia. Debe ir después de corregir_especialistas_no_marcados."""
    mask = df['condicion_laboral_1'].isin(CONDICIONES_EN_FORMACION) & (df['es_especialista'] == 'SI')
    n = int(mask.sum())
    if n != ESPECIALISTAS_EN_FORMACION_ESPERADOS:
        raise ValueError(
            f"quitar_especialidad_en_formacion: {n} filas "
            f"(esperadas {ESPECIALISTAS_EN_FORMACION_ESPERADOS}); revisar antes de continuar"
        )
    df['flag_especialidad_removida_por_formacion'] = mask
    df.loc[mask, 'es_especialista'] = 'NO'
    df.loc[mask, ['id_especialidad', 'especialidad', 'condicion_especialidad']] = pd.NA
    df.loc[mask, 'flag_especialista_incompleto'] = False
    print(f"quitar_especialidad_en_formacion: {n} filas pasan a es_especialista = NO")
    return df

def rellenar_sin_dato_residentes(df):
    """Un residente siempre tiene especialidad en formación; si falta, es un dato
    que debería existir: se marca SIN_DATO para que se vea en Power BI."""
    es_res = _es_residente(df)
    for col in ['especialidad', 'id_especialidad']:
        mask = es_res & df[col].isna()
        df.loc[mask, col] = SIN_DATO_ESPECIALIDAD
        print(f"rellenar_sin_dato_residentes: {col} -> SIN_DATO en {int(mask.sum())}")
    return df

def aplicar_reglas_negocio(df):
    """Marca con flags los registros que violan las reglas de negocio de es_especialista,
    tal como vienen de la fuente (antes de corregir_residentes e igualar_condicion_especialidad,
    que los corrigen). No modifica datos: solo agrega columnas flag_*."""
    campos_incompletos = (
        df['id_especialidad'].isna()
        | df['especialidad'].isna()
        | df['condicion_especialidad'].isna()
    )
    es_si = df['es_especialista'] == 'SI'
    es_residente = _es_residente(df)        # antes: df['condicion_laboral']...str.lower() == 'residente'

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
    consolidado = unificar_regimen_laboral(consolidado)
    consolidado = corregir_cargo(consolidado)
    consolidado = unificar_grupos(consolidado)
    consolidado = agregar_grupofinal4_y_subgrupo(consolidado)
    consolidado = agregar_condiciones_laborales(consolidado)
    consolidado = completar_ids_especialidad(consolidado) 
    consolidado = rellenar_sin_dato(consolidado)
    consolidado = corregir_distritos(consolidado)
    consolidado = corregir_quintil(consolidado)
    consolidado = unificar_unidades_ejecutoras(consolidado)
    consolidado = unificar_establecimientos(consolidado)
    consolidado = agregar_categoria_actual(consolidado)  
    consolidado = limpiar_especialidad(consolidado) 
    consolidado = corregir_ids_sin_nombre(consolidado)
    consolidado = unificar_especialidades(consolidado) 
    consolidado = aplicar_reglas_negocio(consolidado)  
    consolidado = corregir_residentes(consolidado)  
    consolidado = igualar_condicion_especialidad(consolidado)
    consolidado = corregir_especialistas_no_marcados(consolidado) 
    consolidado = quitar_especialidad_tecnicos_auxiliares(consolidado) 
    consolidado = quitar_especialidad_en_formacion(consolidado)              
    consolidado = rellenar_sin_dato_residentes(consolidado)    
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