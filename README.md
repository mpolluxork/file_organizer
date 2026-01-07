# 📧 Outlook Daily Digest

Automatización para generar y enviar un resumen diario de correos de Outlook, organizados por prioridad y con limpieza automática de conversaciones.

## 🎯 Características

- **Resumen HTML elegante** de todos los correos del día
- **Clasificación automática** en "Equipo Directo (VIP)" y "Otros Prioritarios"
- **Deduplicación de conversaciones** - muestra solo el mensaje más reciente de cada hilo
- **Filtrado inteligente** - excluye correos no deseados (spam, newsletters, etc.)
- **Limpieza automática** (opcional) - elimina mensajes antiguos de conversaciones
- **Envío automático** a tu correo personal para revisión fuera de la oficina

## 📋 Requisitos

- Windows con Microsoft Outlook instalado
- Python 3.7 o superior
- Cuenta de Outlook configurada localmente

## 🚀 Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone <repository-url>
   cd file_organizer
   ```

2. **Crear entorno virtual (recomendado):**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configurar el archivo de configuración:**
   - Copiar `config.example.json` a `config.json`
   - Editar `config.json` con tus datos:
     ```json
     {
       "gmail_recipient": "tu-email@gmail.com",
       "equipo_directo": [
         "jefe@empresa.com",
         "colega@empresa.com"
       ],
       "dominios_filtrar": [
         "@spam.com",
         "@newsletter.com"
       ],
       "remitentes_filtrar": [
         "noreply@example.com"
       ],
       "asuntos_filtrar": [
         "Automated:",
         "Newsletter"
       ]
     }
     ```

## 💻 Uso

### Modo Interactivo
```bash
python outlook_daily_digest.py
```
El script te preguntará si deseas eliminar correos antiguos.

### Modo Automático (para tareas programadas)

**Solo generar resumen (sin borrar):**
```bash
python outlook_daily_digest.py --no-delete
```

**Generar resumen y limpiar correos:**
```bash
python outlook_daily_digest.py --delete
```

### Ejecutar con Batch File
```bash
run_outlook_digest.bat
```

## 📅 Programar Ejecución Diaria

Usa el **Programador de Tareas de Windows**:

1. Abrir "Programador de tareas"
2. Crear tarea básica
3. Nombre: "Outlook Daily Digest"
4. Desencadenador: Diariamente a las 18:00 (o tu hora preferida)
5. Acción: Iniciar programa
   - Programa: `C:\ruta\a\tu\venv\Scripts\python.exe`
   - Argumentos: `outlook_daily_digest.py --delete`
   - Iniciar en: `C:\ruta\a\file_organizer`

## 📖 Documentación Adicional

- Ver [`README_DAILY_DIGEST.md`](README_DAILY_DIGEST.md) para detalles técnicos completos

## 🗂️ Estructura del Proyecto

```
file_organizer/
├── outlook_daily_digest.py      # Script principal
├── run_outlook_digest.bat       # Ejecutor batch
├── config.example.json          # Plantilla de configuración
├── requirements.txt             # Dependencias Python
├── README.md                    # Este archivo
├── README_DAILY_DIGEST.md       # Documentación técnica
└── archive/                     # Scripts antiguos (no en Git)
```

## ⚙️ Configuración

El archivo `config.json` permite personalizar:

- **gmail_recipient**: Email donde recibirás el resumen
- **equipo_directo**: Lista de emails VIP (tu equipo directo)
- **dominios_filtrar**: Dominios a excluir del resumen
- **remitentes_filtrar**: Remitentes específicos a excluir
- **asuntos_filtrar**: Patrones en asuntos a excluir

## 🔒 Seguridad

- `config.json` está en `.gitignore` para proteger datos sensibles
- Usa `config.example.json` como plantilla al compartir el código

## 📝 Licencia

Uso personal/interno.

## 🤝 Contribuciones

Este es un proyecto personal, pero sugerencias son bienvenidas.
