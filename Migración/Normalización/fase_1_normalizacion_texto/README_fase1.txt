FASE 1 - NORMALIZACION ESTRUCTURAL BASICA
==================================================

Objetivo:
Estandarizar columnas textuales generales sin alterar aun correos, telefonos, URLs ni logica de deduplicacion.

Reglas aplicadas:
- Trim de espacios al inicio y final
- Colapsar espacios repetidos
- Eliminar caracteres invisibles/control
- Normalizacion Unicode
- Remover tildes
- Convertir a mayusculas
- Limpiar separadores sobrantes en extremos

Archivos generados:
- Base normalizada: /Users/elizabethmesa/ibp/Migración/Normalización/Scripts/fase_1_normalizacion_texto/GlobalDataUpdated-30-3-2026_fase1.xlsx
- Reporte de cambios: /Users/elizabethmesa/ibp/Migración/Normalización/Scripts/fase_1_normalizacion_texto/reporte_cambios_fase1.xlsx
- Resumen de fase: /Users/elizabethmesa/ibp/Migración/Normalización/Scripts/fase_1_normalizacion_texto/README_fase1.txt

Importante:
Esta fase no modifica la base original y no aplica todavia reglas semanticas como quitar SAS, LTDA, INC, ni estandarizar correos, telefonos o LinkedIn.
