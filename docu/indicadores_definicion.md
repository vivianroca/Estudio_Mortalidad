# Definición de indicadores trabajados

Los nueve indicadores cuentan con archivos CSV en formato largo en `data/processed/` e informes de validación en `reports/estructura/`.

| Indicador | Contenido | Cobertura | Estado | Alias |
|---|---|---|---|---|
| **IND-002** | Departamento de ocurrencia y residencia | 2009–2024 | Validado | `ocurrencia_resid` |
| **IND-004** | Residencia, sexo y causa de muerte | 2009–2024 | Códigos finales asignados; advertencia del original de 2023 documentada | `residencia_causa` |
| **IND-005** | Residencia, sexo y edades simples | 2009–2024 | Validado; menores de un año como `0` | `residencia_edad_sexo` |
| **IND-006** | Residencia, sexo y grupos de edad | 2009–2024 | Validado | `residencia_grupo` |
| **IND-007** | Edad, sexo y causa de muerte | 2009–2024, **sin 2011** | Validado; códigos finales asignados | `edad_causa_sexo` |
| **IND-009** | Lugar de ocurrencia por departamento y municipio de residencia | 2009–2024 | Validado | `residencia_lugar` |
| **IND-010** | Mes y día de ocurrencia | 2009–2024 | Validado | `ocurrencia_mes_dia` |
| **IND-013** | Tipo de certificación por departamento y municipio de residencia | 2009–2024 | Validado | `residencia_certifica` |
| **IND-014** | Asistencia recibida por departamento y municipio de residencia | 2009–2024 | Validado | `residencia_asiste` |

Todos los alias son únicos, usan letras sin tildes y guiones bajos, y tienen como máximo 20 caracteres.
