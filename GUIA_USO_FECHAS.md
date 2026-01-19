# Guía Rápida: Procesamiento de Correos por Rango de Fechas

## Problema Resuelto

El script ahora puede procesar correos de días específicos o rangos de fechas, no solo "hoy". Esto te permite recuperar los correos del 14 y 15 de enero que no se procesaron.

## Comandos para Recuperar Correos Perdidos

### Opción 1: Procesar cada día por separado

```powershell
# Procesar solo el 14 de enero
python outlook_daily_digest.py --date 2026-01-14 --delete

# Procesar solo el 15 de enero
python outlook_daily_digest.py --date 2026-01-15 --delete
```

### Opción 2: Procesar ambos días en un solo resumen

```powershell
# Procesar del 14 al 15 de enero en un solo reporte
python outlook_daily_digest.py --start-date 2026-01-14 --end-date 2026-01-15 --delete
```

## Todos los Comandos Disponibles

### Uso Normal (Hoy)

```powershell
# Modo interactivo (te pregunta si quieres borrar)
python outlook_daily_digest.py

# Procesar hoy CON borrado
python outlook_daily_digest.py --delete

# Procesar hoy SIN borrado (solo resumen)
python outlook_daily_digest.py --no-delete
```

### Uso con Fechas Específicas

```powershell
# Un día específico
python outlook_daily_digest.py --date 2026-01-14 --delete

# Rango de fechas
python outlook_daily_digest.py --start-date 2026-01-14 --end-date 2026-01-15 --delete

# Solo resumen, sin borrar (cualquier fecha)
python outlook_daily_digest.py --date 2026-01-14 --no-delete
```

## Qué Hace el Script

1. **Filtra correos** del rango de fechas especificado
2. **Agrupa conversaciones** (muestra solo el más reciente de cada hilo)
3. **Clasifica** en "Equipo Directo (VIP)" y "Otros Prioritarios"
4. **Marca como leídos** todos los correos procesados
5. **Elimina** (si usas `--delete`):
   - Correos filtrados (UTEL, LinkedIn, etc.)
   - Mensajes antiguos de conversaciones (excepto los que tienen adjuntos)
6. **Genera reporte HTML** y lo envía a tu Gmail

## Recomendación para Tu Caso

Para recuperar los correos del 14 y 15 de enero, te recomiendo:

```powershell
python outlook_daily_digest.py --start-date 2026-01-14 --end-date 2026-01-15 --delete
```

Esto procesará ambos días en un solo resumen y marcará todos los correos como leídos.

## Notas Importantes

- El formato de fecha es **YYYY-MM-DD** (año-mes-día)
- Si usas `--delete`, los correos se marcarán como leídos Y se eliminarán los antiguos
- Si usas `--no-delete`, solo se genera el resumen (nada se borra ni marca como leído)
- El resumen se envía a tu Gmail configurado en `config.json`
