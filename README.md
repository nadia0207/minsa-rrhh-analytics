# MINSA RRHH Analytics

Proyecto de análisis de datos de personal de salud del sector público peruano (MINSA), 
usando bases de datos abiertas de dotación de RRHH correspondientes al periodo 2019-2025.

🚧🚧🚧🚧 **Proyecto en construcción** — ver progreso en los commits.

## Objetivo

Analizar la evolución de la dotación de personal de salud a nivel nacional (por región, 
categoría ocupacional y régimen laboral) y construir un modelo predictivo sobre [pendiente 
de definir: rotación / proyección de dotación].

## Stack

- **Python** (pandas, SQLAlchemy) — limpieza y consolidación de datos (ETL)
- **PostgreSQL** — almacenamiento y consultas analíticas (vistas, KPIs)
- **Power BI** — dashboard interactivo conectado directo a la base de datos
- **Scikit-learn** — modelo de Machine Learning
- **uv** — gestión de entorno y dependencias

## Fuente de datos

Bases de dotación de personal de salud (diciembre de cada año, 2019-2025), de acceso 
público, publicadas por el Ministerio de Salud del Perú (MINSA). 
https://digep.minsa.gob.pe/bdatos.html

> Nota: los archivos originales no se incluyen en este repositorio por su tamaño 
> (>200,000 filas por año). Ver `data/raw/README.md` para instrucciones de descarga.

## Estructura del proyecto

```
minsa-rrhh-analytics/
├── data/
│   ├── raw/            # Excel originales (no versionados)
│   └── processed/      # Datos limpios y consolidados
├── notebooks/          # Exploración y análisis
├── src/minsa_analystics/
│   ├── etl.py           # Carga y consolidación de los 7 años
│   ├── db.py             # Conexión a PostgreSQL
│   └── models.py         # Modelo de ML
├── sql/                  # Vistas y consultas analíticas
└── README.md
```

## Progreso

- [x] Estructura del proyecto y entorno con uv
- [x] ETL: consolidación de los 7 años (2019-2025) en un solo dataset de ~1.9 M de filas
- [x] Comparación de columnas entre años y mapeo a un esquema común (`notebooks/01_comparacion_columnas.ipynb`)
- [ ] Validación de calidad del consolidado (`notebooks/02_validacion_consolidado.ipynb`) — **en curso**
  - [x] Normalización de `es_especialista` (SI/NO) y unificación de tipos
  - [ ] Reglas de negocio (especialidad completa, residentes)
  - [ ] Revisión de edad
- [ ] Carga a PostgreSQL
- [ ] Análisis exploratorio (EDA)
- [ ] Dashboard en Power BI
- [ ] Modelo de Machine Learning

## Decisiones de calidad de datos

- Las bases de cada año usan nombres y codificaciones distintas; se unificaron en un esquema común y se documentaron las columnas descartadas y el motivo.
- `es_especialista` venía con codificaciones distintas según el año (`SI`/`NO`, `No`, `0`, `-`, vacío). Se normaliza a `SI`/`NO`: los valores vacíos, `0` y `-` se tratan como `NO`.
- El ETL se detiene con un error si aparece un valor no reconocido, en lugar de clasificarlo en silencio.

## Autora

Nadia Llamoca Córdova — [LinkedIn](https://linkedin.com/in/nadiallamoca) · [GitHub](https://github.com/nadia0207)