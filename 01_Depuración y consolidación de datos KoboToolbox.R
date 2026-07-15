# =============================================================================
# Proyecto: Estudio de carga de VSR en Colombia
# Script:   01 - Depuración y consolidación de datos KoboToolbox
# Autora:   Rosmery Vanessa Barroso Parra
# Fecha:    Junio 2026
# =============================================================================


# Librerías -------------------------------------------------------------------

library(tidyverse)
library(lubridate)
library(janitor)
library(openxlsx)


# Ruta del archivo ------------------------------------------------------------

# Actualizar con la ruta y nombre del archivo Excel exportado desde KoboToolbox
RUTA_EXCEL <- "Kobo6.xlsx"


# Importación de hojas --------------------------------------------------------

# Hoja principal (nivel paciente)
datos_main <- read.xlsx(RUTA_EXCEL, sheet = 1, detectDates = TRUE) %>%
  clean_names()

# Repeat group: eventos VSR (uno o más por paciente)
datos_eventos <- read.xlsx(RUTA_EXCEL, sheet = "seq_event", detectDates = TRUE) %>%
  clean_names()

# Repeat groups de nivel de atención por evento
datos_urgencias       <- read.xlsx(RUTA_EXCEL, sheet = "grupo_urgencias",       detectDates = TRUE) %>% clean_names()
datos_hospitalizacion <- read.xlsx(RUTA_EXCEL, sheet = "grupo_hospitalizacion", detectDates = TRUE) %>% clean_names()
datos_cuidados_inter  <- read.xlsx(RUTA_EXCEL, sheet = "grupo_cintermedios",    detectDates = TRUE) %>% clean_names()
datos_uci             <- read.xlsx(RUTA_EXCEL, sheet = "grupo_uci",             detectDates = TRUE) %>% clean_names()

# Repeat groups de uso de recursos por evento
datos_imagenes        <- read.xlsx(RUTA_EXCEL, sheet = "seq_n_imagen",          detectDates = TRUE) %>% clean_names()
datos_procedimientos  <- read.xlsx(RUTA_EXCEL, sheet = "seq_n_procedimientos",  detectDates = TRUE) %>% clean_names()
datos_interconsultas  <- read.xlsx(RUTA_EXCEL, sheet = "seq_n_interconsultas",  detectDates = TRUE) %>% clean_names()
datos_medicamentos    <- read.xlsx(RUTA_EXCEL, sheet = "seq_n_tto",             detectDates = TRUE) %>% clean_names()


# Muestra asignada y metas por médico -----------------------------------------

# Ajustar n_muestra según el protocolo de muestreo de cada centro
muestra_asignada <- tribble(
  ~centro,             ~n_muestra,
  "clinica_rosario",   450,
  "colsubsidio",       1839,
  "erasmo",            334,
  "HINF",              677
)

# Ajustar cod_medico, centro y meta_semanal con los datos reales del estudio
medicos_meta <- tribble(
  ~cod_medico,  ~centro,             ~meta_semanal,
  "SLMC",       "erasmo",            31,   # Sebastian Moncada
  "NB",         "erasmo",            31,   # Neyker Bautista
  "SSL2",       "erasmo",            31,   # ¿variante/typo de SLMC?
  "SL",         "erasmo",            31,   # ¿variante/typo de SLMC?
  "ET",         "HINF",              50,   # Elloth Tamara
  "MCGM",       "HINF",              25,   # Maria Claudia Guerra
  "KVCA",       "HINF",              30,   # Karen Carvajal
  "kvca",       "HINF",              30,   # idem minúscula
  "CI",         "clinica_rosario",   13,   # Cielo Isaza
  "MC",         "clinica_rosario",   10,   # Marlly Correa
  "DBP",        "colsubsidio",       31,   # Daniela Ballen
  "LPV",        "colsubsidio",       13,   # Liliana Prieto
  "MDLAJC",     "colsubsidio",       25,    # ¿María Jimenez?,
  "ACOV",      "clinica_rosario",    10,    #Ana Cristina Osorio
  "EMG",       "colsubsidio",        25,    #Esteban Montoya
  "MVSF",     "colsubsidio",         25    #María Verónica Suárez
)


# Fecha de corte del informe
fecha_corte <- as.Date("2026-07-12")
fecha_semana_anterior <- fecha_corte - 6

# Limpieza general de la hoja principal ---------------------------------------

main_limpio <- datos_main %>%
  mutate(
    # Fechas
    fecha_nacimiento = as.Date(fecha_nacimiento),
    start          = as.Date(trunc(end), origin = "1899-12-30"),
    end            = as.Date(trunc(end), origin = "1899-12-30"),
    fecha_registro = as.Date(trunc(end), origin = "1899-12-30"),
    fecha_nacimiento = as.Date(trunc(fecha_nacimiento), origin = "1899-12-30"),
    cod_medico         = str_squish(toupper(cod_medico)),
    nombre_institucion = str_squish(nombre_institucion),
    # # Flag de elegibilidad
    elegible = (criterio_fecha     == "si" &
                  criterio_edad      == "si" &
                  criterio_clinico   == "si" &
                  criterio_exclusion == "si")
  )


# Unificación de variantes, nombres reales, minúsculas y exclusión de registros a borrar
main_limpio <- main_limpio %>%
  filter(!cod_medico %in% c("DBP-BORRAR", "EMG-BORRAR", "AG", "ag", "ACOV-BORRAR")) %>%
  mutate(
    cod_medico = case_when(
      cod_medico %in% c("SSL2", "SL") ~ "SLMC",
      cod_medico == "kvca"            ~ "KVCA",
      TRUE                            ~ cod_medico
    )
  ) %>%
  mutate(
    nombre_medico = case_when(
      cod_medico == "SLMC"   ~ "Sebastián Moncada",
      cod_medico == "ET"     ~ "Elloth Tamara",
      cod_medico == "MCGM"   ~ "Maria Claudia",
      cod_medico == "CI"     ~ "Cielo",
      cod_medico == "MC"     ~ "Marlly",
      cod_medico == "KVCA"   ~ "Karen Carvajal",
      cod_medico == "NB"     ~ "Neyker Bautista",
      cod_medico == "DBP"    ~ "Daniela Ballen",
      cod_medico == "LPV"    ~ "Liliana Prieto",
      cod_medico == "MDLAJC" ~ "María Jimenez",
      cod_medico == "ACOV"    ~ "Ana Cristina Osorio",
      cod_medico == "EMG"    ~ "Esteban Montoya",
      cod_medico == "MVSF" ~ "María Verónica Suárez",
      TRUE                   ~ cod_medico
    )
  )

# Limpieza de eventos ---------------------------------------------------------

eventos_limpio <- datos_eventos %>%
  mutate(
    fecha_inicio_evento = as.Date(fecha_inicio_evento),
    fecha_fin_evento    = as.Date(fecha_fin_evento),
    dias_evento_calc    = as.numeric(fecha_fin_evento - fecha_inicio_evento) + 1
  )


# Depuración: checks de calidad -----------------------------------------------

# --- check 1: pacientes duplicados por centro ---------------------------------
duplicados <- main_limpio %>%
  group_by(nombre_institucion, cod_participante) %>%
  filter(n() > 1) %>%
  arrange(nombre_institucion, cod_participante, fecha_registro) %>%
  mutate(flag = "Paciente duplicado en el mismo centro")

# --- check 2: fecha fin de evento anterior a fecha inicio --------------------
errores_fechas_evento <- eventos_limpio %>%
  filter(!is.na(fecha_fin_evento) & !is.na(fecha_inicio_evento)) %>%
  filter(fecha_fin_evento < fecha_inicio_evento) %>%
  mutate(flag = "Fecha fin anterior a fecha inicio del evento")

# --- check 3: días calculados no coinciden con los reportados por Kobo -------
errores_dias_evento <- eventos_limpio %>%
  filter(!is.na(dias_evento) & !is.na(dias_evento_calc)) %>%
  mutate(diferencia_dias = dias_evento_calc - as.numeric(dias_evento)) %>%
  filter(abs(diferencia_dias) > 0) %>%
  mutate(flag = "Discrepancia entre días calculados y reportados")

# --- check 4: urgencias declaradas sin fechas registradas --------------------
ids_con_urgencias <- datos_urgencias %>% pull(submission_id) %>% unique()

errores_urgencias_sin_fechas <- eventos_limpio %>%
  filter(urgencias == "si") %>%
  filter(!submission_id %in% ids_con_urgencias) %>%
  mutate(flag = "Declara urgencias pero no hay fechas registradas")

# --- check 5: eventos sin ningún medicamento registrado ----------------------
ids_con_medicamentos <- datos_medicamentos %>% pull(parent_index) %>% unique()

errores_sin_medicamentos <- eventos_limpio %>%
  filter(!index %in% ids_con_medicamentos) %>%
  mutate(flag = "Evento sin medicamento registrado")

# --- check 6: campos 'otro' vacíos -------------------------------------------
errores_otro_vacio <- main_limpio %>%
  select(id, nombre_institucion, cod_participante, cod_medico,
         matches("^otra|^otras|_otro$")) %>%
  pivot_longer(-c(id, nombre_institucion, cod_participante, cod_medico),
               names_to = "variable", values_to = "valor") %>%
  filter(is.na(valor) | str_trim(as.character(valor)) == "") %>%
  mutate(flag = "Campo 'otro' vacío sin especificación")

# Cálculo permanente de edad al inicio del evento -----------------------------

eventos_limpio <- eventos_limpio %>%
  left_join(main_limpio %>% select(id, fecha_nacimiento), by = c("submission_id" = "id")) %>%
  mutate(
    edad_dias_evento  = as.numeric(difftime(fecha_inicio_evento, fecha_nacimiento, units = "days")),
    edad_meses_evento = round(edad_dias_evento / 30.4375, 1)
  )

# --- check 7: edad al ingreso fuera de rango (0-59 meses) -------------------
errores_edad <- eventos_limpio %>%
  left_join(
    main_limpio %>%
      select(id, cod_participante, cod_medico, nombre_medico, nombre_institucion),
    by = c("submission_id" = "id")
  ) %>%
  filter(!is.na(edad_meses_evento)) %>%
  filter(edad_meses_evento < 0 | edad_meses_evento > 59) %>%
  mutate(flag = "Edad fuera del rango de elegibilidad (0-59 meses)") %>%
  select(
    submission_id, index, cod_participante, cod_medico, nombre_medico,
    nombre_institucion, fecha_nacimiento, fecha_inicio_evento,
    edad_meses_evento, flag
  )

# --- check 8: control de imágenes/procedimientos/interconsultas --------------
# Si marcó si en el control, debe tener al menos 1 registro en el repeat group

ids_con_imagenes <- datos_imagenes %>% pull(parent_index) %>% unique()
errores_imagenes <- eventos_limpio %>%
  filter(control_imagenes_lab == "si") %>%
  filter(!index %in% ids_con_imagenes) %>%
  mutate(flag = "Declara imágenes/labs pero no hay registros")

ids_con_procedimientos <- datos_procedimientos %>% pull(parent_index) %>% unique()
errores_procedimientos <- eventos_limpio %>%
  filter(control_procedimientos == "si") %>%
  filter(!index %in% ids_con_procedimientos) %>%
  mutate(flag = "Declara procedimientos pero no hay registros")

ids_con_interconsultas <- datos_interconsultas %>% pull(parent_index) %>% unique()
errores_interconsultas <- eventos_limpio %>%
  filter(atencion_internconsulta == "si") %>%
  filter(!index %in% ids_con_interconsultas) %>%
  mutate(flag = "Declara interconsultas pero no hay registros")


# Tablas de seguimiento -------------------------------------------------------

# Pacientes únicos elegibles por centro
pacientes_por_centro <- main_limpio %>%
  filter(elegible == TRUE, fecha_registro <= fecha_corte) %>%
  group_by(nombre_institucion) %>%
  summarise(n_pacientes_kobo = n_distinct(id), .groups = "drop")

eventos_por_centro <- eventos_limpio %>%
  left_join(
    main_limpio %>% 
      filter(elegible == TRUE, fecha_registro <= fecha_corte) %>% 
      select(id, nombre_institucion),
    by = c("submission_id" = "id")
  ) %>%
  filter(!is.na(nombre_institucion)) %>%
  group_by(nombre_institucion) %>%
  summarise(n_eventos_kobo = n(), .groups = "drop")

# Tabla 1 - resumen por centro
tabla_centros <- muestra_asignada %>%
  left_join(pacientes_por_centro, by = c("centro" = "nombre_institucion")) %>%
  left_join(eventos_por_centro,   by = c("centro" = "nombre_institucion")) %>%
  mutate(
    n_pacientes_kobo = replace_na(n_pacientes_kobo, 0),
    n_eventos_kobo   = replace_na(n_eventos_kobo, 0),
    pct_pacientes    = round(n_pacientes_kobo / n_muestra * 100, 1)
  )


total_centros <- tabla_centros %>%
  summarise(
    centro           = "TOTAL",
    n_muestra        = sum(n_muestra),
    n_pacientes_kobo = sum(n_pacientes_kobo),
    n_eventos_kobo   = sum(n_eventos_kobo),
    pct_pacientes    = round(sum(n_pacientes_kobo) / sum(n_muestra) * 100, 1)
  )

tabla_centros_final <- bind_rows(tabla_centros, total_centros)

# Seguimiento por médico
inicio_semana_actual <- floor_date(fecha_corte, unit = "week", week_start = 1)

seguimiento_medico_full <- main_limpio %>%
  filter(elegible == TRUE, fecha_registro <= fecha_corte) %>%
  group_by(cod_medico) %>%
  summarise(
    total_acumulado = n_distinct(id),
    ultima_semana   = n_distinct(id[fecha_registro >= inicio_semana_actual]),
    .groups = "drop"
  ) %>%
  left_join(medicos_meta, by = "cod_medico") %>%
  mutate(
    pct_cumplimiento_semana    = round(ultima_semana   / meta_semanal * 100, 1),
    pct_cumplimiento_acumulado = round(total_acumulado / meta_semanal * 100, 1)
  )


# Errores fecha nacimiento --------------------------------------------------------
# --- check 8: fecha de nacimiento posterior al inicio del evento -------------

errores_fecha_nacimiento <- eventos_limpio %>%
  left_join(
    main_limpio %>%
      select(
        id,
        cod_participante,
        cod_medico,
        nombre_medico,
        nombre_institucion
      ),
    by = c("submission_id" = "id")
  ) %>%
  filter(
    !is.na(fecha_nacimiento),
    !is.na(fecha_inicio_evento),
    fecha_nacimiento > fecha_inicio_evento
  ) %>%
  mutate(flag = "Fecha de nacimiento posterior al inicio del evento") %>%
  select(
    submission_id, index, cod_participante, cod_medico, nombre_medico,
    nombre_institucion, fecha_nacimiento, fecha_inicio_evento, flag
  )

errores_fecha_nacimiento %>%
  count(cod_medico, nombre_medico, sort = TRUE)


# Tabla 4 - Máximo nivel de atención por episodio, por médico y centro --------

# Identificar qué episodios tienen registro en cada nivel
ids_urgencias       <- datos_urgencias       %>% pull(parent_index) %>% unique()
ids_hospitalizacion <- datos_hospitalizacion %>% pull(parent_index) %>% unique()
ids_cuidados_inter  <- datos_cuidados_inter  %>% pull(parent_index) %>% unique()
ids_uci             <- datos_uci             %>% pull(parent_index) %>% unique()

# Clasificar cada episodio por su máximo nivel (jerárquico)
episodios_nivel <- eventos_limpio %>%
  left_join(
    main_limpio %>%
      filter(elegible == TRUE, fecha_registro <= fecha_corte) %>%
      select(id, cod_medico, nombre_medico, nombre_institucion),
    by = c("submission_id" = "id")
  ) %>%
  filter(!is.na(nombre_institucion)) %>%
  mutate(
    parent_index_num = as.numeric(parent_index),   # por si acaso hay diferencia de tipo
    max_nivel = case_when(
      parent_index_num %in% ids_uci             ~ "UCI",
      parent_index_num %in% ids_cuidados_inter  ~ "Cuidados intermedios",
      parent_index_num %in% ids_hospitalizacion ~ "Hospitalización general",
      parent_index_num %in% ids_urgencias       ~ "Urgencias",
      TRUE                                      ~ "Sin nivel registrado"
    ),
    max_nivel = factor(max_nivel, levels = c(
      "Urgencias", "Hospitalización general",
      "Cuidados intermedios", "UCI", "Sin nivel registrado"
    ))
  )

# --- check 9: eventos elegibles sin ningún nivel de atención registrado ------
# El médico se guía por submission_id (id de Kobo) directamente, no por
# cod_participante, así que estos objetos no incluyen esa columna.

errores_sin_estancia <- episodios_nivel %>%
  filter(max_nivel == "Sin nivel registrado") %>%
  mutate(flag = "Evento elegible sin ningún nivel de atención registrado") %>%
  select(
    submission_id, index, cod_medico, nombre_medico,
    nombre_institucion, fecha_inicio_evento, fecha_fin_evento, flag
  )

# Pacientes elegibles cuyos eventos están TODOS sin nivel registrado
# (es decir, quedarían totalmente fuera de la base de atenciones)

pacientes_sin_estancia <- episodios_nivel %>%
  group_by(submission_id, cod_medico, nombre_medico, nombre_institucion) %>%
  summarise(
    n_eventos_totales      = n(),
    n_eventos_sin_estancia = sum(max_nivel == "Sin nivel registrado"),
    .groups = "drop"
  ) %>%
  filter(n_eventos_totales == n_eventos_sin_estancia)

message(glue::glue(
  "Eventos sin ningún nivel de atención registrado: {nrow(errores_sin_estancia)} | ",
  "Pacientes que quedarían totalmente sin atención registrada: {nrow(pacientes_sin_estancia)}"
))

# Agregar: n y % por médico, centro y nivel
tabla_nivel_medico <- episodios_nivel %>%
  group_by(nombre_institucion, cod_medico, nombre_medico, max_nivel) %>%
  summarise(n = n(), .groups = "drop") %>%
  group_by(nombre_institucion, cod_medico) %>%
  mutate(
    total_medico = sum(n),
    pct          = round(n / total_medico * 100, 1)
  ) %>%
  ungroup() %>%
  pivot_wider(
    names_from  = max_nivel,
    values_from = c(n, pct),
    values_fill = 0,
    names_glue  = "{max_nivel}_{.value}"    # e.g. "Urgencias_n", "Urgencias_pct"
  ) %>%
  # Reordenar columnas: primero identificadores, luego pares n/pct por nivel
  select(
    nombre_institucion, cod_medico, nombre_medico, total_medico,
    starts_with("Urgencias"),
    starts_with("Hospitalización general"),
    starts_with("Cuidados intermedios"),
    starts_with("UCI"),
    starts_with("Sin nivel")
  )

# Fila TOTAL (colapsa médicos)
total_nivel <- episodios_nivel %>%
  group_by(max_nivel) %>%
  summarise(n = n(), .groups = "drop") %>%
  mutate(
    total_medico       = sum(n),
    pct                = round(n / total_medico * 100, 1),
    nombre_institucion = "TOTAL",
    cod_medico         = "",
    nombre_medico      = "TOTAL"
  ) %>%
  pivot_wider(
    names_from  = max_nivel,
    values_from = c(n, pct),
    values_fill = 0,
    names_glue  = "{max_nivel}_{.value}"
  ) %>%
  select(names(tabla_nivel_medico))     

tabla_nivel_medico_final <- bind_rows(tabla_nivel_medico, total_nivel)


# Tabla 3 - resumen de errores -------------------------------------------------
# Va después de episodios_nivel/errores_sin_estancia porque depende de ellos

resumen_errores <- tibble(
  tipo_error = c(
    "Pacientes duplicados",
    "Fechas de evento inconsistentes (fin < inicio)",
    "Días calculados vs. reportados discrepantes",
    "Urgencias declaradas sin fechas registradas",
    "Eventos sin medicamento registrado",
    "Campos 'otro' vacíos sin especificación",
    "Edad fuera de rango de elegibilidad (0-59 meses)",
    "Imágenes/labs declarados sin registros",
    "Procedimientos declarados sin registros",
    "Interconsultas declaradas sin registros",
    "Eventos sin ningún nivel de atención registrado",
    "Pacientes totalmente sin atención registrada"
  ),
  n_registros_afectados = c(
    nrow(duplicados),
    nrow(errores_fechas_evento),
    nrow(errores_dias_evento),
    nrow(errores_urgencias_sin_fechas),
    nrow(errores_sin_medicamentos),
    nrow(errores_otro_vacio),
    nrow(errores_edad),
    nrow(errores_imagenes),
    nrow(errores_procedimientos),
    nrow(errores_interconsultas),
    nrow(errores_sin_estancia),
    nrow(pacientes_sin_estancia)
  )
) %>%
  mutate(estado = if_else(n_registros_afectados == 0, "Sin errores", "Requiere revisión"))


# Exportación -----------------------------------------------------------------

save(
  tabla_centros_final,
  seguimiento_medico_full,
  resumen_errores,
  duplicados,
  errores_fechas_evento,
  errores_dias_evento,
  errores_urgencias_sin_fechas,
  errores_sin_medicamentos,
  errores_otro_vacio,
  errores_edad,
  errores_fecha_nacimiento,
  errores_imagenes,
  errores_procedimientos,
  errores_interconsultas,
  errores_sin_estancia,      # <-- nuevo
  pacientes_sin_estancia,    # <-- nuevo
  main_limpio,
  eventos_limpio,
  datos_urgencias,
  datos_hospitalizacion,
  datos_cuidados_inter,
  datos_uci,
  datos_imagenes,
  datos_procedimientos,
  datos_interconsultas,
  datos_medicamentos,
  medicos_meta,
  muestra_asignada,
  tabla_nivel_medico_final,
  episodios_nivel,
  file = "datos_depurados_VSR.RData"
)

message("Depuración completada. Objetos guardados en 'datos_depurados_VSR.RData'")