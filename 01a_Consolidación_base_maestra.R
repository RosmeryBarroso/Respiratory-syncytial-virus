# =============================================================================
# Proyecto: Estudio de carga de VSR en Colombia
# Script:   02 - Conformación de la base maestra de atenciones
# Autora:   Rosmery Vanessa Barroso Parra
# Fecha:    Julio 2026
# Objetivo: A partir de los objetos depurados en el script 01
#           (datos_depurados_VSR.RData), consolidar las cuatro hojas de
#           estancia en la base maestra de atenciones (una fila por evento),
#           y dejar listas las bases auxiliares de medicamentos,
#           procedimientos, interconsultas e imágenes, con las columnas de
#           estandarización creadas para que la Dra. Daniela las complete.
#
# Insumo:   datos_depurados_VSR.RData  (generado por 01_depuracion_kobo.R)
# Salida:   base_maestra_VSR.RData
#           base_atenciones_VSR.xlsx
#           valores_unicos_para_estandarizar.xlsx
# =============================================================================


# Librerías -------------------------------------------------------------------

library(tidyverse)
library(lubridate)
library(janitor)
library(openxlsx)


# Carga de insumos --------------------------------------------------------------

load("datos_depurados_VSR.RData")

eventos_limpio <- eventos_limpio %>%
  mutate(index = as.numeric(index))

# fecha_corte no quedó guardada en el .RData del script 01: se redefine acá
# con el mismo valor. Ajustar si cambia la fecha de corte del informe.
fecha_corte <- as.Date("2026-07-12")


# 1. verificación de criterios de inclusión del estudio ------------------------

# La elegibilidad clínica ya fue validada en el script 01 y quedó en
# main_limpio$elegible. Acá solo se filtra y se deja constancia del universo.

pacientes_elegibles <- main_limpio %>%
  filter(elegible == TRUE, fecha_registro <= fecha_corte)

n_elegibles <- n_distinct(pacientes_elegibles$id)

message(glue::glue("Pacientes elegibles: {n_elegibles} (meta: 3,300)"))

if (n_elegibles != 3300) {
  warning(glue::glue(
    "El número de pacientes elegibles ({n_elegibles}) no coincide con la meta de 3,300."
  ))
}


# 2. estandarización de las hojas de estancia -----------------------------------

# Las 4 hojas comparten esquema: fecha_inicio_x, fecha_egreso_x, dias_x,
# parent_index (enlaza con eventos_limpio$index).

estandarizar_estancia <- function(df, nivel) {
  df %>%
    select(
      parent_index,
      submission_id,
      fecha_inicio_reportada = matches("^fecha_inicio_"),
      fecha_egreso_reportada = matches("^fecha_egreso_"),
      dias_reportados_kobo   = matches("^dias_[a-z]+$")   # excluye *_calculo y *_error
    ) %>%
    mutate(
      fecha_inicio_reportada = as.Date(fecha_inicio_reportada),
      fecha_egreso_reportada = as.Date(fecha_egreso_reportada),
      nivel_atencion = nivel
    )
}

urgencias_std       <- estandarizar_estancia(datos_urgencias,       "Urgencias")
hospitalizacion_std <- estandarizar_estancia(datos_hospitalizacion, "Hospitalización general")
intermedios_std     <- estandarizar_estancia(datos_cuidados_inter,  "Cuidados intermedios")
uci_std             <- estandarizar_estancia(datos_uci,             "UCI")


# 3. recálculo de duraciones y fechas por nivel de atención ---------------------

# Un mismo evento puede tener varias filas en el mismo nivel (traslados).
# Por eso se recalcula fila a fila y luego se agrega por evento sumando los
# días y tomando el rango de fechas.

recalcular_nivel <- function(df) {
  
  df_fila <- df %>%
    mutate(
      fecha_inicio_reportada = as.Date(trunc(fecha_inicio_reportada), origin = "1899-12-30"),
      fecha_egreso_reportada = as.Date(trunc(fecha_egreso_reportada), origin = "1899-12-30"),
      dias_fila_calc = as.numeric(fecha_egreso_reportada - fecha_inicio_reportada) + 1,
      dias_reportados_kobo = as.numeric(dias_reportados_kobo),
      diferencia_dias = dias_fila_calc - dias_reportados_kobo,
      flag_discrepancia = !is.na(diferencia_dias) & diferencia_dias != 0,
      flag_fecha_invertida = !is.na(fecha_egreso_reportada) & !is.na(fecha_inicio_reportada) &
        fecha_egreso_reportada < fecha_inicio_reportada
    )
  
  df_evento <- df_fila %>%
    group_by(parent_index, submission_id, nivel_atencion) %>%
    summarise(
      fecha_inicio_nivel  = suppressWarnings(min(fecha_inicio_reportada, na.rm = TRUE)),
      fecha_egreso_nivel  = suppressWarnings(max(fecha_egreso_reportada, na.rm = TRUE)),
      dias_nivel_calc     = sum(dias_fila_calc, na.rm = TRUE),
      n_registros_nivel   = n(),
      n_discrepancias     = sum(flag_discrepancia, na.rm = TRUE),
      n_fechas_invertidas = sum(flag_fecha_invertida, na.rm = TRUE),
      n_sin_fecha_inicio  = sum(is.na(fecha_inicio_reportada)),
      n_sin_fecha_egreso  = sum(is.na(fecha_egreso_reportada)),
      .groups = "drop"
    ) %>%
    mutate(
      fecha_inicio_nivel = if_else(is.infinite(fecha_inicio_nivel), as.Date(NA), fecha_inicio_nivel),
      fecha_egreso_nivel = if_else(is.infinite(fecha_egreso_nivel), as.Date(NA), fecha_egreso_nivel)
    )
  
  list(detalle = df_fila, evento = df_evento)
}

urgencias_rec       <- recalcular_nivel(urgencias_std)
hospitalizacion_rec <- recalcular_nivel(hospitalizacion_std)
intermedios_rec     <- recalcular_nivel(intermedios_std)
uci_rec             <- recalcular_nivel(uci_std)

calidad_recalculo_estancias <- bind_rows(
  urgencias_rec$evento, hospitalizacion_rec$evento,
  intermedios_rec$evento, uci_rec$evento
) %>%
  filter(n_discrepancias > 0 | n_fechas_invertidas > 0 |
           n_sin_fecha_inicio == n_registros_nivel | n_sin_fecha_egreso == n_registros_nivel) %>%
  mutate(flag = case_when(
    n_fechas_invertidas > 0 ~ "Fecha de egreso anterior a ingreso",
    n_discrepancias > 0     ~ "Días calculados difieren de lo reportado en Kobo",
    n_sin_fecha_inicio == n_registros_nivel ~ "Sin ninguna fecha de inicio registrada en el nivel",
    n_sin_fecha_egreso == n_registros_nivel ~ "Sin ninguna fecha de egreso registrada en el nivel",
    TRUE ~ NA_character_
  ))


# 4. consolidación de estancias en la base de atenciones -------------------------

niveles_evento <- bind_rows(
  urgencias_rec$evento,
  hospitalizacion_rec$evento,
  intermedios_rec$evento,
  uci_rec$evento
)

orden_niveles <- c("Urgencias", "Hospitalización general",
                   "Cuidados intermedios", "UCI")

base_atenciones <- niveles_evento %>%
  mutate(nivel_atencion = factor(nivel_atencion, levels = orden_niveles)) %>%
  group_by(parent_index, submission_id) %>%
  summarise(
    fecha_ingreso_atencion = min(fecha_inicio_nivel, na.rm = TRUE),
    fecha_egreso_atencion  = max(fecha_egreso_nivel, na.rm = TRUE),
    dias_totales_estancia  = sum(dias_nivel_calc, na.rm = TRUE),
    max_nivel_atencion     = orden_niveles[max(as.integer(nivel_atencion))],
    niveles_transitados    = paste(sort(unique(as.character(nivel_atencion))), collapse = " -> "),
    n_niveles              = n_distinct(nivel_atencion),
    .groups = "drop"
  ) %>%
  mutate(
    duracion_calendario = as.numeric(fecha_egreso_atencion - fecha_ingreso_atencion) + 1
  )

base_atenciones <- base_atenciones %>%
  left_join(
    eventos_limpio %>%
      select(index_2, submission_id, fecha_inicio_evento, fecha_fin_evento, edad_dias_evento, edad_meses_evento),
    by = c("parent_index" = "index_2", "submission_id")
  ) %>%
  left_join(
    pacientes_elegibles %>%
      select(id, cod_participante, cod_medico, nombre_medico, nombre_institucion,
             fecha_nacimiento, sexo, departamento, municipio, estrato,
             tipo_afiliacion, educacion),
    by = c("submission_id" = "id")
  ) %>%
  filter(!is.na(nombre_institucion)) %>%
  relocate(cod_participante, nombre_institucion, cod_medico, nombre_medico,
           .after = submission_id)

# eventos elegibles sin ninguna estancia registrada en las 4 hojas
eventos_elegibles <- eventos_limpio %>%
  inner_join(pacientes_elegibles %>% select(id), by = c("submission_id" = "id"))

eventos_sin_estancia <- eventos_elegibles %>%
  anti_join(base_atenciones, by = c("index_2" = "parent_index", "submission_id")) %>%
  mutate(flag = "Evento elegible sin ninguna estancia registrada en las 4 hojas")

message(glue::glue(
  "Eventos elegibles: {nrow(eventos_elegibles)} | ",
  "Eventos en base_atenciones: {nrow(base_atenciones)} | ",
  "Eventos sin estancia: {nrow(eventos_sin_estancia)}"
))


# 5. cálculo de costos por nivel de atención (pendiente tarifario) ---------------

# Cuando llegue el tarifario, se une por max_nivel_atencion (o por nivel_atencion
# a nivel de niveles_evento, si el costo se calcula por cada nivel transitado):
#
# tarifario <- read.xlsx("tarifario.xlsx") %>% clean_names()
#
# base_atenciones <- base_atenciones %>%
#   left_join(tarifario, by = c("max_nivel_atencion" = "nivel_tarifario")) %>%
#   mutate(costo_estimado = dias_totales_estancia * valor_dia)


# 6. consolidación de bases auxiliares: medicamentos, procedimientos, interconsultas e imágenes ----

# Cada base auxiliar queda enlazada al evento (parent_index/submission_id) y
# con las columnas de estandarización ya creadas, vacías, para que la
# Dra. Daniela las complete con el catálogo de equivalencias y SISMED.

medicamentos_aux <- datos_medicamentos %>%
  left_join(eventos_limpio %>% select(index_2, submission_id), by = c("parent_index" = "index_2")) %>%
  mutate(
    grupo_farmacologico_estandarizado = NA_character_,
    nombre_generico_estandarizado     = NA_character_,
    precio_unitario_sismed            = NA_real_,
    costo_total_medicamento           = NA_real_
  )

procedimientos_aux <- datos_procedimientos %>%
  left_join(eventos_limpio %>% select(index_2, submission_id), by = c("parent_index" = "index_2")) %>%
  mutate(
    servicio_estandarizado = NA_character_,
    valor_procedimiento    = NA_real_
  )

interconsultas_aux <- datos_interconsultas %>%
  left_join(eventos_limpio %>% select(index_2, submission_id), by = c("parent_index" = "index_2")) %>%
  mutate(
    especialidad_estandarizada = NA_character_,
    valor_interconsulta        = NA_real_
  )

imagenes_aux <- datos_imagenes %>%
  left_join(eventos_limpio %>% select(index_2, submission_id), by = c("parent_index" = "index_2")) %>%
  mutate(
    servicio_estandarizado = NA_character_,
    valor_examen            = NA_real_
  )

# 7. valores únicos para estandarización con la Dra. Daniela ---------------------

# Listado de nombres tal como quedaron digitados en Kobo, para que sirvan de
# insumo al catálogo de equivalencias. Se separan el campo principal y el
# texto libre de "otro", porque cada uno se revisa distinto.

unicos_medicamentos_grupo <- datos_medicamentos %>%
  select(tto_grupo, tto_broncodilatador, tto_antibioticos, tto_cinhalado,
         tto_csistemico, tto_antipiretico, tto_spray_nasal, tto_fluidos,
         tto_jarabe, tto_otro) %>%
  distinct() %>%
  arrange(tto_grupo)

unicos_medicamentos_otro <- datos_medicamentos %>%
  select(tto_bronco_otro, tto_antibiotico_otro, tto_cinhalado_otro,
         tto_csistemico_otro, tto_antipiretico_otro, tto_spray_nasal_otro,
         tto_fluidos_otro, tto_otro_otro) %>%
  mutate(across(everything(), as.character)) %>%
  pivot_longer(everything(), names_to = "campo_otro", values_to = "texto_libre") %>%
  filter(!is.na(texto_libre), str_trim(texto_libre) != "") %>%
  distinct() %>%
  arrange(campo_otro, texto_libre)

unicos_procedimientos <- datos_procedimientos %>%
  select(tipo_procedimiento, procedimiento_quirurgico, procedimiento_no_quirurgico) %>%
  distinct() %>%
  arrange(tipo_procedimiento)

unicos_procedimientos_otro <- datos_procedimientos %>%
  select(procedimiento_otro_quirurgico, procedimiento_otro_noquirurgico) %>%
  mutate(across(everything(), as.character)) %>%
  pivot_longer(everything(), names_to = "campo_otro", values_to = "texto_libre") %>%
  filter(!is.na(texto_libre), str_trim(texto_libre) != "") %>%
  distinct() %>%
  arrange(campo_otro, texto_libre)

unicos_interconsultas <- datos_interconsultas %>%
  select(tipo_servicio) %>%
  distinct() %>%
  arrange(tipo_servicio)

unicos_interconsultas_otro <- datos_interconsultas %>%
  select(otro_profesional) %>%
  filter(!is.na(otro_profesional), str_trim(otro_profesional) != "") %>%
  distinct() %>%
  arrange(otro_profesional)

unicos_imagenes <- datos_imagenes %>%
  select(laboratorio_nombre) %>%
  distinct() %>%
  arrange(laboratorio_nombre)

unicos_imagenes_otro <- datos_imagenes %>%
  select(laboratorio_otro) %>%
  filter(!is.na(laboratorio_otro), str_trim(laboratorio_otro) != "") %>%
  distinct() %>%
  arrange(laboratorio_otro)


# 8. exportación -------------------------------------------------------------

save(
  base_atenciones,
  niveles_evento,
  calidad_recalculo_estancias,
  eventos_sin_estancia,
  medicamentos_aux,
  procedimientos_aux,
  interconsultas_aux,
  imagenes_aux,
  pacientes_elegibles,
  file = "base_maestra_VSR.RData"
)

write.xlsx(
  list(
    "base_atenciones"      = base_atenciones,
    "calidad_recalculo"    = calidad_recalculo_estancias,
    "eventos_sin_estancia" = eventos_sin_estancia,
    "medicamentos"         = medicamentos_aux,
    "procedimientos"       = procedimientos_aux,
    "interconsultas"       = interconsultas_aux,
    "imagenes"             = imagenes_aux
  ),
  file = "base_atenciones_VSR.xlsx",
  overwrite = TRUE
)

write.xlsx(
  list(
    "medicamentos_grupo"      = unicos_medicamentos_grupo,
    "medicamentos_otro"       = unicos_medicamentos_otro,
    "procedimientos"          = unicos_procedimientos,
    "procedimientos_otro"     = unicos_procedimientos_otro,
    "interconsultas"          = unicos_interconsultas,
    "interconsultas_otro"     = unicos_interconsultas_otro,
    "imagenes"                = unicos_imagenes,
    "imagenes_otro"           = unicos_imagenes_otro
  ),
  file = "valores_unicos_para_estandarizar.xlsx",
  overwrite = TRUE
)

message("Conformación de base maestra completada.")