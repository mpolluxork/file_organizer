# Outlook Daily Digest - Resumen Diario de Correos

Sistema automático para generar y enviar resúmenes diarios de correos de Outlook a Gmail.

## 🎯 Características

### Clasificación Inteligente
- ✅ **Usa la clasificación nativa de Outlook**: "Prioritarios" (Focused) vs "Otros" (Other)
- ✅ **Solo muestra correos prioritarios** en el resumen detallado
- ✅ **Conversaciones inteligentes**: Agrupa correos por hilo y muestra solo el más reciente

### Categorización VIP
- ⭐ **Equipo Directo (VIP)**: Destacado con preview de contenido (200 caracteres)


- 📬 **Otros Prioritarios**: Solo asunto y remitente (sin preview)

### Resumen HTML
- 📊 Estadísticas en la parte superior
- 🎨 Diseño moderno y responsive
- 📧 Se envía automáticamente a: **tucorreo@gmail.com**

### Limpieza Automática (Opcional)
- 🗑️ Puede eliminar correos clasificados como "Otros" después de generar el resumen
- ⚠️ Requiere confirmación explícita

## 🚀 Uso

### Ejecución Manual

```bash
python outlook_daily_digest.py
```

El script te preguntará:

```
¿Deseas eliminar los correos clasificados como 'Otros' (no prioritarios)?
  1. No, solo generar resumen
  2. Sí, eliminar correos 'Otros' después de generar resumen

Selecciona opción (1 o 2):
```

### Automatización con Task Scheduler (Futuro)

Para ejecutar automáticamente cada noche a las 23:00:

1. Abrir **Task Scheduler** (Programador de tareas)
2. Crear tarea básica
3. Configurar:
   - Nombre: "Outlook Daily Digest"
   - Trigger: Diario a las 23:00
   - Acción: Ejecutar programa
   - Programa: `python`
   - Argumentos: `outlook_daily_digest.py`
   - Directorio: `C:\file_organizer`

## 📋 Qué hace el script

1. **Lee correos de hoy** (desde las 00:00 hasta ahora)
2. **Filtra por "Prioritarios"** usando la clasificación de Outlook
3. **Deduplica conversaciones** - solo muestra el mensaje más reciente de cada hilo
4. **Categoriza**:
   - Equipo Directo (VIP)
   - Otros Prioritarios
   - Otros (solo cuenta, no muestra detalles)
5. **Genera HTML** con diseño moderno
6. **Envía a Gmail** automáticamente

## 📊 Ejemplo de Resumen

```
┌─────────────────────────────────────────┐
│     📧 Resumen Diario de Correos        │
│            05/01/2026                   │
├─────────────────────────────────────────┤
│  5          12         45        18     │
│ Equipo    Otros     Otros    Conver-    │
│ Directo  Priority  (filtro)  saciones   │
├─────────────────────────────────────────┤
│ ⭐ Equipo Directo (VIP)                 │
│                                         │
│ • Juanito Perez                         │
│   RE: PROYECTOS 2026       │
│   Preview: Hola Buenos tardes...        │
│                                         │
├─────────────────────────────────────────┤
│ 📬 Otros Prioritarios                   │
│                                         │
│ • Microsoft Security (2 correos)        │
│   - Seminario web | Proteja sus datos  │
│   - Alertas de empleo                   │
│                                         │
└─────────────────────────────────────────┘
```

## 🔧 Configuración

Para cambiar el destinatario del resumen, edita la línea 17 en `outlook_daily_digest.py`:

```python
GMAIL_RECIPIENT = "tucorreo@gmail.com"
```

Para modificar la lista de Equipo Directo, edita las líneas 20-27:

```python
EQUIPO_DIRECTO = [
    'juanitoperez@tuempresa.mx',
    # ... agregar o quitar emails
]
```

## 🐛 Troubleshooting

### El script no detecta "Focused Inbox"
- Asegúrate de tener activada la Bandeja Prioritaria en Outlook
- El script asume que los correos sin clasificación son "Focused" por defecto

### No se envía el correo
- Verifica que Outlook esté abierto y conectado
- Revisa que la cuenta de Outlook tenga permisos para enviar correos

### Errores de conexión COM
- Cierra y vuelve a abrir Outlook
- Ejecuta el script con permisos de administrador

## 📝 Notas

- El script solo procesa correos recibidos **hoy** (desde las 00:00)
- Las conversaciones se deduplicán automáticamente
- Los correos "Otros" se cuentan pero no se muestran en detalle
- La eliminación de "Otros" es opcional y requiere confirmación

## 🔮 Mejoras Futuras

- [ ] Integración con Ollama para resúmenes con IA
- [ ] Marcar correos como leídos después de incluirlos en el resumen
- [ ] Configuración vía archivo JSON
- [ ] Soporte para múltiples destinatarios
- [ ] Resúmenes semanales/mensuales
