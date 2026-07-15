# ============================================================
# CONSOLIDADO DE MUESTRAS ESTRATIFICADAS:VSR
# Distribución de la muestra y duración del evento, por centro
# ============================================================

library(dplyr)
library(readxl)
library(tidyr)
library(stringr)
library(flextable)

# ------------------------------------------------------------
# Rutas de las 4 muestras estratificadas (ajustar nombres de archivo)
# ------------------------------------------------------------
rutas <- c(
  "Colsubsidio"         = "datos/Muestras finales/MUESTRA_ESTRATIFICADA_Colsubsidio.xlsx",
  "Clínica del Rosario" = "datos/Muestras finales/MUESTRA_ESTRATIFICADA_Rosario.xlsx",
  "HNFP"                = "datos/Muestras finales/MUESTRA_ESTRATIFICADA_HINFP.xlsx",
  "Erasmo Meoz"         = "datos/Muestras finales/MUESTRA_ESTRATIFICADA_ERASMO.xlsx"
)

# ------------------------------------------------------------
# Unifica las etiquetas de subgrupo (ITRI en Colsubsidio/HNFP/Erasmo,
# RIPS en Rosario) a tres categorías comparables entre centros
# ------------------------------------------------------------
normalizar_subgrupo <- function(x) {
  case_when(
    str_starts(x, "Grupo 1") ~ "Grupo 1",
    str_starts(x, "Grupo 2") ~ "Grupo 2",
    str_starts(x, "Grupo 3") ~ "Grupo 3",
    TRUE ~ NA_character_
  )
}

# ------------------------------------------------------------
# Lector genérico de cada muestra: cada centro trae columnas propias
# (id vs ingreso, resultado vs resultado_vsr, columnas extra como
# complejidad_laboratrios o es_nuevo). Aquí solo se conservan las
# columnas comunes necesarias para las tablas.
# ------------------------------------------------------------
leer_muestra <- function(ruta, centro) {
  datos <- read_excel(ruta, sheet = "MUESTRA ESTRATIFICADA")
  
  # llave de paciente: 'id' en Colsubsidio/HNFP/Erasmo, 'ingreso' en Rosario
  if ("id" %in% names(datos)) {
    datos <- datos %>% rename(id_paciente = id)
  } else {
    datos <- datos %>% rename(id_paciente = ingreso)
  }
  
  # id_kobo puede no existir en todos los archivos (por si acaso)
  if (!"id_kobo" %in% names(datos)) {
    datos <- datos %>% mutate(id_kobo = NA_integer_)
  }
  
  # color_grupo: clasificación original, antes del recálculo de subgrupo
  if (!"color_grupo" %in% names(datos)) {
    datos <- datos %>% mutate(color_grupo = NA_character_)
  }
  
  datos %>%
    mutate(
      centro                  = centro,
      subgrupo_norm           = normalizar_subgrupo(subgrupo),
      subgrupo_norm_original  = normalizar_subgrupo(color_grupo),
      fecha_inicial_evento    = as.Date(fecha_inicial_evento, format = "%d/%m/%Y"),
      fecha_final_evento      = as.Date(fecha_final_evento,   format = "%d/%m/%Y")
    ) %>%
    select(centro, id_paciente, id_kobo, evento_vsr, subgrupo_norm, subgrupo_norm_original,
           fecha_inicial_evento, fecha_final_evento)
}

# ------------------------------------------------------------
# Cargar y unir los 4 centros
# ------------------------------------------------------------
muestra_consolidada <- bind_rows(
  mapply(leer_muestra, rutas, names(rutas), SIMPLIFY = FALSE)
)
# Una fila por evento VSR (evita duplicar por múltiples filas de
# historia clínica que pertenecen al mismo evento)
eventos_unicos <- muestra_consolidada %>%
  distinct(centro, id_paciente, id_kobo, evento_vsr, subgrupo_norm, subgrupo_norm_original,
           fecha_inicial_evento, fecha_final_evento)


cat("Total eventos únicos en la muestra consolidada:", nrow(eventos_unicos), "\n")
cat("Total pacientes únicos:",
    n_distinct(paste(eventos_unicos$centro, eventos_unicos$id_paciente)), "\n")
cat("Eventos por subgrupo (normalizado):\n")
print(table(eventos_unicos$subgrupo_norm, useNA = "ifany"))

cat("Distribución de color_grupo (clasificación original) por centro:\n")
print(table(eventos_unicos$centro, eventos_unicos$subgrupo_norm_original, useNA = "ifany"))



# DIAGNÓSTICO 1: cuota de diseño (max id_kobo) vs pacientes únicos


verificacion_centro <- eventos_unicos %>%
  group_by(centro) %>%
  summarise(
    pacientes_unicos = n_distinct(id_paciente),
    cuota_diseno      = max(id_kobo, na.rm = TRUE),
    n_id_kobo_unicos  = n_distinct(id_kobo),
    diferencia        = pacientes_unicos - cuota_diseno,
    .groups = "drop"
  )

print(verificacion_centro)


# DIAGNÓSTICO 2: ¿hay huecos en la secuencia de id_kobo?
# (si cuota_diseno = 300 pero solo hay 298 valores únicos de id_kobo,
#  significa que faltan números en la secuencia 1..300)


huecos_id_kobo <- eventos_unicos %>%
  group_by(centro) %>%
  summarise(
    max_kobo     = max(id_kobo, na.rm = TRUE),
    secuencia_ok  = n_distinct(id_kobo) == max(id_kobo, na.rm = TRUE),
    faltantes     = list(setdiff(1:max(id_kobo, na.rm = TRUE), unique(id_kobo))),
    .groups = "drop"
  )

print(huecos_id_kobo)



# DIAGNÓSTICO 3: ¿hay un mismo id_paciente con dos id_kobo distintos?
# (esto explicaría por qué pacientes_unicos < cuota_diseno:
#  dos "cupos" de la muestra fueron asignados por error al mismo paciente)


pacientes_con_dos_kobo <- eventos_unicos %>%
  distinct(centro, id_paciente, id_kobo) %>%
  group_by(centro, id_paciente) %>%
  filter(n_distinct(id_kobo) > 1) %>%
  arrange(centro, id_paciente)

print(pacientes_con_dos_kobo)


# ============================================================
# Totales generales del estudio (para % de la fila "Total")
# ============================================================
total_pacientes_general <- n_distinct(eventos_unicos$id_paciente)
total_eventos_general   <- n_distinct(paste(eventos_unicos$id_paciente, eventos_unicos$evento_vsr))



# ============================================================
# TABLA A:Distribución de la muestra por centro
# (pacientes, eventos y % del total, todos los subgrupos incluidos)
# ============================================================

tabla_a_datos <- eventos_unicos %>%
  group_by(centro) %>%
  summarise(
    pacientes = n_distinct(id_paciente),
    eventos   = n_distinct(paste(id_paciente, evento_vsr)),
    .groups   = "drop"
  ) %>%
  bind_rows(
    summarise(., centro = "Total", pacientes = sum(pacientes), eventos = sum(eventos))
  ) %>%
  mutate(
    n_pac = pacientes[centro == "Total"],
    n_ev  = eventos[centro == "Total"],
    `# Pacientes` = ifelse(
      centro == "Total", as.character(pacientes),
      paste0(pacientes, " (", round(pacientes / n_pac * 100, 1), "%)")
    ),
    `# Eventos` = ifelse(
      centro == "Total", as.character(eventos),
      paste0(eventos, " (", round(eventos / n_ev * 100, 1), "%)")
    )
  ) %>%
  select(Centro = centro, `# Pacientes`, `# Eventos`)

tabla_a <- tabla_a_datos %>%
  flextable() %>%
  autofit() %>%
  theme_box() %>%
  bold(i = ~ Centro == "Total") %>%
  bg(bg = "#6EA1BA", part = "header") %>%
  color(color = "white", part = "header") %>%
  bold(part = "header") %>%
  align(align = "center", part = "all") %>%
  align(j = 1, align = "left", part = "body") %>%
  add_header_lines("Distribución de la muestra por centro") %>%
  bg(bg = "#D9E8F2", part = "header", i = 1) %>%
  color(color = "black", part = "header", i = 1)

tabla_a


# ============================================================
# Totales generales del estudio (para % de la fila "Total")
# ============================================================

total_pacientes_general <- n_distinct(eventos_unicos$id_paciente)
total_eventos_general   <- n_distinct(paste(eventos_unicos$id_paciente, eventos_unicos$evento_vsr))


# ============================================================
# Clasificación de la duración del evento en meses
# ============================================================

clasificar_duracion <- function(dias) {
  meses <- dias / 30
  case_when(
    meses <= 1              ~ "\u2264 1 mes",
    meses > 1 & meses <= 2  ~ "1 a 2 meses",
    meses > 2 & meses <= 3  ~ "2 a 3 meses",
    meses > 3               ~ "> 3 meses"
  )
}

orden_bins <- c("\u2264 1 mes", "1 a 2 meses", "2 a 3 meses", "> 3 meses")

# ------------------------------------------------------------
# Tabla de distribución de duración por centro, CON porcentaje
# (por fila de centro: % respecto al total de esa fila;
#  fila "Total": % respecto al total del grupo/tabla, no del estudio general)
# ------------------------------------------------------------
tabla_duracion_por_centro <- function(data, grupos_incluidos, titulo) {
  
  base_duracion <- data %>%
    filter(subgrupo_norm %in% grupos_incluidos) %>%
    mutate(
      dias_evento  = as.numeric(fecha_final_evento - fecha_inicial_evento),
      bin_duracion = clasificar_duracion(dias_evento),
      bin_duracion = factor(bin_duracion, levels = orden_bins)
    )
  
  tabla_wide <- base_duracion %>%
    count(centro, bin_duracion) %>%
    complete(centro, bin_duracion, fill = list(n = 0)) %>%
    pivot_wider(names_from = bin_duracion, values_from = n)
  
  fila_total <- tabla_wide %>%
    summarise(centro = "Total", across(where(is.numeric), sum))
  
  tabla_conteos <- bind_rows(tabla_wide, fila_total) %>%
    mutate(Total = rowSums(across(where(is.numeric))))
  
  # total del grupo/tabla (no del estudio general)
  total_ev_grupo <- tabla_conteos$Total[tabla_conteos$centro == "Total"]
  
  formatear_pct <- function(valor, base) {
    ifelse(base == 0, as.character(valor),
           paste0(valor, " (", round(valor / base * 100, 1), "%)"))
  }
  
  tabla_final <- tabla_conteos %>%
    rowwise() %>%
    mutate(across(
      all_of(orden_bins),
      ~ if (centro == "Total") formatear_pct(.x, total_ev_grupo) else formatear_pct(.x, Total)
    )) %>%
    ungroup() %>%
    mutate(Total = ifelse(
      centro == "Total",
      paste0(Total, " (", round(Total / total_ev_grupo * 100, 1), "%)"),
      as.character(Total)
    )) %>%
    rename(Centro = centro)
  
  tabla_final %>%
    flextable() %>%
    autofit() %>%
    theme_box() %>%
    bold(i = ~ Centro == "Total") %>%
    bg(bg = "#6EA1BA", part = "header") %>%
    color(color = "white", part = "header") %>%
    bold(part = "header") %>%
    align(align = "center", part = "all") %>%
    align(j = 1, align = "left", part = "body") %>%
    add_header_lines(titulo) %>%
    bg(bg = "#D9E8F2", part = "header", i = 1) %>%
    color(color = "black", part = "header", i = 1)
}



# ============================================================
# TABLA B (ORIGINAL): Grupo 1 según color_grupo, antes del recálculo
# ============================================================
eventos_unicos_original <- eventos_unicos %>%
  mutate(subgrupo_norm = subgrupo_norm_original)

tabla_b_original <- tabla_duracion_por_centro(
  eventos_unicos_original,
  grupos_incluidos = "Grupo 1",
  titulo = "Distribución de la duración del evento: Grupo 1 (clasificación original, antes del recálculo)"
)

tabla_b_original

tabla_grupo1_original <- tabla_conteo_grupo(
  eventos_unicos_original,
  grupo  = "Grupo 1",
  titulo = "Grupo 1 (clasificación original, antes del recálculo) - conteo de pacientes y eventos"
)

tabla_grupo1_original

# ============================================================
# TABLA B:Duración del evento, Grupo 1 (cruzaron con RIPS/ITRI)
# ============================================================

tabla_b <- tabla_duracion_por_centro(
  eventos_unicos,
  grupos_incluidos = "Grupo 1",
  titulo = "Distribución de la duración del evento:Grupo 1 (con cruce)"
)

tabla_b


# ------------------------------------------------------------
# Tabla de conteo (pacientes/eventos) por centro, CON porcentaje
# (por fila de centro: % respecto al total de ese grupo;
#  fila "Total": % respecto al total del grupo, no del estudio general)
# ------------------------------------------------------------
tabla_conteo_grupo <- function(data, grupo, titulo) {
  
  datos <- data %>%
    filter(subgrupo_norm == grupo) %>%
    group_by(centro) %>%
    summarise(
      pacientes = n_distinct(id_paciente),
      eventos   = n_distinct(paste(id_paciente, evento_vsr)),
      .groups   = "drop"
    )
  
  fila_total <- datos %>%
    summarise(centro = "Total", pacientes = sum(pacientes), eventos = sum(eventos))
  
  bind_rows(datos, fila_total) %>%
    mutate(
      n_pac_grupo = pacientes[centro == "Total"],
      n_ev_grupo  = eventos[centro == "Total"],
      `# Pacientes` = paste0(pacientes, " (", round(pacientes / n_pac_grupo * 100, 1), "%)"),
      `# Eventos`   = paste0(eventos, " (", round(eventos / n_ev_grupo * 100, 1), "%)")
    ) %>%
    select(Centro = centro, `# Pacientes`, `# Eventos`) %>%
    flextable() %>%
    autofit() %>%
    theme_box() %>%
    bold(i = ~ Centro == "Total") %>%
    bg(bg = "#6EA1BA", part = "header") %>%
    color(color = "white", part = "header") %>%
    bold(part = "header") %>%
    align(align = "center", part = "all") %>%
    align(j = 1, align = "left", part = "body") %>%
    add_header_lines(titulo) %>%
    bg(bg = "#D9E8F2", part = "header", i = 1) %>%
    color(color = "black", part = "header", i = 1)
}


# ============================================================
# GRUPO 2 y GRUPO 3
# ============================================================

tabla_grupo2 <- tabla_conteo_grupo(
  eventos_unicos,
  grupo  = "Grupo 2",
  titulo = "Grupo 2:Fuera de ventana ITRI (solo conteo, sin distribución de duración)"
)

tabla_grupo2

tabla_grupo3 <- tabla_conteo_grupo(
  eventos_unicos,
  grupo  = "Grupo 3",
  titulo = "Grupo 3:Sin ingreso registrado (solo conteo, sin distribución de duración)"
)

tabla_grupo3