"""
Outlook Daily Digest - Resumen Diario de Correos
Genera un resumen HTML de los correos del día y lo envía por email.
"""

import win32com.client
from datetime import datetime, timedelta
from collections import defaultdict
import sys
import json
import os
import html


class OutlookDailyDigest:
    """Generador de resumen diario de correos de Outlook."""
    
    def __init__(self):
        """Inicializa la clase cargando configuración y conexión con Outlook."""
        self.load_config()
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
        except Exception as e:
            print(f"❌ Error al conectar con Outlook: {e}")
            sys.exit(1)

    def load_config(self):
        """Carga la configuración desde config.json."""
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        
        # Valores por defecto
        self.GMAIL_RECIPIENT = ""
        self.EQUIPO_DIRECTO = []
        self.DOMINIOS_FILTRAR = []
        self.REMITENTES_FILTRAR = []
        self.ASUNTOS_FILTRAR = []

        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.GMAIL_RECIPIENT = config.get('gmail_recipient', "")
                    self.EQUIPO_DIRECTO = config.get('equipo_directo', [])
                    self.DOMINIOS_FILTRAR = config.get('dominios_filtrar', [])
                    self.REMITENTES_FILTRAR = config.get('remitentes_filtrar', [])
                    self.ASUNTOS_FILTRAR = config.get('asuntos_filtrar', [])
                print("✓ Configuración cargada desde config.json")
            except Exception as e:
                print(f"⚠ Error al cargar config.json: {e}. Usando valores por defecto.")
        else:
            print("⚠ Archivo config.json no encontrado. Usando valores por defecto.")
    
    def get_inbox(self):
        """Obtiene la carpeta de bandeja de entrada."""
        try:
            return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        except Exception as e:
            print(f"✗ Error al acceder a la bandeja de entrada: {e}")
            sys.exit(1)
    
    def get_sender_email(self, message):
        """Extrae el email del remitente."""
        try:
            if message.SenderEmailType == 'EX':
                sender = message.Sender
                if sender:
                    try:
                        return sender.GetExchangeUser().PrimarySmtpAddress.lower()
                    except:
                        return message.SenderEmailAddress.lower()
            else:
                return message.SenderEmailAddress.lower()
        except:
            return "unknown@unknown.com"
    
    def should_filter_email(self, sender_email, subject):
        """
        Determina si un correo debe ser filtrado (excluido del resumen).
        
        Args:
            sender_email: Email del remitente
            subject: Asunto del correo
        
        Returns:
            True si debe filtrarse, False si debe incluirse
        """
        sender_email_lower = sender_email.lower()
        subject_lower = subject.lower()
        
        # Filtrar por dominio
        for dominio in self.DOMINIOS_FILTRAR:
            if dominio.lower() in sender_email_lower:
                return True
        
        # Filtrar por remitente específico
        if sender_email_lower in [r.lower() for r in self.REMITENTES_FILTRAR]:
            return True
        
        # Filtrar por patrones en el asunto
        for patron in self.ASUNTOS_FILTRAR:
            if patron.lower() in subject_lower:
                return True
        
        return False
    
    def is_focused(self, message):
        """
        Detecta si el correo está en la bandeja 'Prioritarios' (Focused Inbox).
        
        NOTA: La propiedad MAPI para Focused Inbox no está disponible en esta versión.
        Por ahora, tratamos TODOS los correos como prioritarios.
        
        Returns:
            True (siempre, hasta que se implemente detección alternativa)
        """
        # TODO: Implementar detección alternativa de Focused Inbox
        # Por ahora, mostrar todos los correos
        return True
    
    def get_todays_emails(self, delete_otros=False, start_date=None, end_date=None):
        """
        Obtiene los correos recibidos en un rango de fechas.
        
        Args:
            delete_otros: Si True, elimina los correos clasificados como "Otros"
            start_date: Fecha inicial (datetime.date). Si es None, usa hoy.
            end_date: Fecha final (datetime.date). Si es None, usa hoy.
        
        Returns:
            Dict con emails categorizados
        """
        inbox = self.get_inbox()
        messages = inbox.Items
        
        # Determinar rango de fechas
        if start_date is None:
            start_date = datetime.now().date()
        if end_date is None:
            end_date = datetime.now().date()
        
        # Validar que start_date <= end_date
        if start_date > end_date:
            print(f"❌ Error: La fecha inicial ({start_date}) no puede ser posterior a la fecha final ({end_date})")
            sys.exit(1)
        
        # Filtrar por rango de fechas
        # Usar el día anterior al start_date para capturar todo el día inicial
        filter_start = start_date - timedelta(days=1)
        # Usar el día siguiente al end_date para capturar todo el día final
        filter_end = end_date + timedelta(days=1)
        
        # Formato: DD/MM/YYYY para Outlook
        filter_string = f"[ReceivedTime] > '{filter_start.strftime('%d/%m/%Y')}' AND [ReceivedTime] < '{filter_end.strftime('%d/%m/%Y')}'"
        messages = messages.Restrict(filter_string)
        messages.Sort("[ReceivedTime]", True)  # Más recientes primero
        
        if start_date == end_date:
            print(f"\n📊 Analizando correos del {start_date.strftime('%Y-%m-%d')}")
        else:
            print(f"\n📊 Analizando correos del {start_date.strftime('%Y-%m-%d')} al {end_date.strftime('%Y-%m-%d')}")
        print(f"   Total de correos encontrados: {messages.Count}")
        
        # Categorizar correos
        equipo_directo_emails = []
        otros_prioritarios = []
        filtered_count = 0  # Correos filtrados (UTEL, LinkedIn, etc.)
        deleted_count = 0   # Correos eliminados
        
        # Para deduplicar conversaciones
        conversations = {}  # conversation_id -> latest_message
        
        total_processed = 0
        
        # CRÍTICO: Convertir a lista para evitar que se salten mensajes al borrar
        # Si iteramos directamente sobre messages e Items y borramos durante la iteración,
        # Outlook modifica la colección y causa que se salten elementos
        messages_list = list(messages)
        
        for message in messages_list:
            try:
                total_processed += 1
                
                # Obtener información del correo
                sender_email = self.get_sender_email(message)
                sender_name = message.SenderName
                subject = message.Subject
                received_time = message.ReceivedTime
                body = message.Body
                is_unread = message.UnRead
                conversation_id = message.ConversationID
                
                # Filtrar solo correos dentro del rango de fechas
                if received_time.date() < start_date or received_time.date() > end_date:
                    continue
                
                # Aplicar filtros personalizados (UTEL, LinkedIn, asuntos, etc.)
                if self.should_filter_email(sender_email, subject):
                    filtered_count += 1
                    if delete_otros:
                        try:
                            message.Delete()
                            deleted_count += 1
                        except Exception as e:
                            print(f"  ⚠ Error eliminando correo filtrado '{subject[:40]}...': {e}")
                    continue
                
                # Solo procesar correos que pasaron los filtros
                email_data = {
                    'sender_email': sender_email,
                    'sender_name': sender_name,
                    'subject': subject,
                    'received_time': received_time,
                    'body': body[:500],  # Primeros 500 caracteres
                    'is_unread': is_unread,
                    'conversation_id': conversation_id,
                    'has_attachments': message.Attachments.Count > 0,
                    'message_object': message  # Guardar referencia al mensaje original
                }
                
                # DEBUG: Mostrar información de conversación
                conv_id_str = conversation_id if conversation_id else "NONE"
                print(f"  📧 [{total_processed}] {subject[:50]}... | ConvID: {conv_id_str[:20]}...")
                
                # Agrupar por conversación (guardar el más reciente para mostrar, pero trackear todos para borrar)
                if conversation_id:
                    if conversation_id not in conversations:
                        # Inicializar con el primer mensaje encontrado
                        conversations[conversation_id] = {
                            'latest_message': email_data,
                            'has_vip': sender_email in self.EQUIPO_DIRECTO,
                            'all_messages': [email_data]  # Para poder borrar los antiguos
                        }
                    else:
                        # Agregar a la lista de todos los mensajes
                        conversations[conversation_id]['all_messages'].append(email_data)
                        print(f"     ↳ Agrupado con conversación existente ({len(conversations[conversation_id]['all_messages'])} mensajes)")
                        
                        # Actualizar si este mensaje es más reciente
                        if received_time > conversations[conversation_id]['latest_message']['received_time']:
                            conversations[conversation_id]['latest_message'] = email_data
                        
                        # Actualizar flag de VIP si este remitente es VIP
                        if sender_email in self.EQUIPO_DIRECTO:
                            conversations[conversation_id]['has_vip'] = True
                else:
                    # Si no tiene conversation_id, tratarlo como único
                    unique_key = f"no_conv_{total_processed}"
                    conversations[unique_key] = {
                        'latest_message': email_data,
                        'has_vip': sender_email in self.EQUIPO_DIRECTO,
                        'all_messages': [email_data]
                    }
                
            except Exception as e:
                print(f"  ⚠ Error procesando correo: {e}")
                continue
        
        # Procesar conversaciones: extraer el mensaje más reciente para mostrar, borrar los antiguos
        conversation_deleted_count = 0
        final_emails = []
        
        for conv_id, conv_data in conversations.items():
            # Extraer el mensaje más reciente (este es el que se mostrará)
            latest_message = conv_data['latest_message']
            
            # Marcar si la conversación tiene VIP
            latest_message['is_vip_conversation'] = conv_data['has_vip']
            
            # Agregar solo el más reciente a la lista final (para evitar anidación)
            final_emails.append(latest_message)
            
            # Borrar mensajes antiguos de la conversación si delete_otros está activado
            all_messages = conv_data['all_messages']
            if len(all_messages) > 1 and delete_otros:
                print(f"\n  🗑️  Procesando conversación con {len(all_messages)} mensajes:")
                # Ordenar por fecha para identificar cuál es el más reciente
                all_messages.sort(key=lambda x: x['received_time'], reverse=True)
                
                # Mostrar todos los mensajes de la conversación
                for idx, msg in enumerate(all_messages):
                    status = "MANTENER (más reciente)" if idx == 0 else f"BORRAR (antiguo, adjuntos={msg['has_attachments']})"
                    print(f"     [{idx+1}] {msg['received_time'].strftime('%H:%M')} - {msg['subject'][:40]}... | {status}")
                
                # Borrar todos excepto el primero (más reciente) y los que tienen adjuntos
                # PERO marcar TODOS como leídos (incluso los que se conservan por adjuntos)
                for old_message in all_messages[1:]:
                    # Marcar como leído SIEMPRE (se borre o no)
                    try:
                        old_message['message_object'].UnRead = False
                    except Exception as e:
                        print(f"     ⚠ Error marcando mensaje como leído: {e}")
                    
                    # Solo borrar si NO tiene adjuntos
                    if not old_message['has_attachments']:
                        try:
                            old_message['message_object'].Delete()
                            conversation_deleted_count += 1
                            print(f"     ✓ Borrado: {old_message['subject'][:40]}...")
                        except Exception as e:
                            print(f"     ⚠ Error eliminando mensaje antiguo de conversación: {e}")
                    else:
                        print(f"     ⊗ Conservado (tiene adjuntos) y marcado como leído: {old_message['subject'][:40]}...")
            
            # Marcar el mensaje más reciente como LEÍDO (ya que aparecerá en el resumen)
            if delete_otros:
                try:
                    latest_message['message_object'].UnRead = False
                except Exception as e:
                    print(f"  ⚠ Error marcando mensaje como leído: {e}")
        
        # Clasificar correos finales por categoría
        equipo_directo_emails = []
        otros_prioritarios = []
        
        for email_data in final_emails:
            # Usar la marca de VIP en conversación
            if email_data.get('is_vip_conversation', False):
                equipo_directo_emails.append(email_data)
            else:
                otros_prioritarios.append(email_data)
        
        # Ordenar por fecha (más recientes primero)
        equipo_directo_emails.sort(key=lambda x: x['received_time'], reverse=True)
        otros_prioritarios.sort(key=lambda x: x['received_time'], reverse=True)
        
        print(f"   Conversaciones únicas: {len(conversations)}")
        print(f"   Equipo Directo: {len(equipo_directo_emails)}")
        print(f"   Otros Prioritarios: {len(otros_prioritarios)}")
        print(f"   🚫 Filtrados (UTEL, LinkedIn, etc.): {filtered_count}")
        
        if delete_otros:
            total_deleted = deleted_count + conversation_deleted_count
            print(f"   🗑️  Correos filtrados eliminados: {deleted_count}")
            print(f"   🗑️  Mensajes antiguos de conversaciones eliminados: {conversation_deleted_count}")
            print(f"   🗑️  Total eliminados: {total_deleted}")
        
        return {
            'equipo_directo': equipo_directo_emails,
            'otros_prioritarios': otros_prioritarios,
            'filtered_count': filtered_count,
            'deleted_count': deleted_count if delete_otros else 0,
            'conversation_deleted_count': conversation_deleted_count if delete_otros else 0,
            'total_conversations': len(conversations),
            'total_processed': total_processed
        }
    
    def generate_html_report(self, email_data, date_range_str=None):
        """
        Genera un reporte HTML bonito con los correos del día o rango de fechas.
        
        Args:
            email_data: Diccionario con correos categorizados
            date_range_str: String con el rango de fechas a mostrar (opcional)
        
        Returns:
            String con HTML del reporte
        """
        if date_range_str is None:
            date_range_str = datetime.now().strftime('%d/%m/%Y')
        
        equipo_directo = email_data['equipo_directo']
        otros_prioritarios = email_data['otros_prioritarios']
        filtered_count = email_data.get('filtered_count', 0)
        total_conversations = email_data['total_conversations']
        
        # Agrupar "Otros Prioritarios" por remitente
        otros_by_sender = defaultdict(list)
        for email in otros_prioritarios:
            otros_by_sender[email['sender_name']].append(email)
        
        html_output = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resumen Diario - {date_range_str}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
            margin: 0;
            padding: 20px;
            color: #333;
        }}
        .container {{
            max-width: 900px;
            margin: 0 auto;
            background-color: white;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 28px;
            font-weight: 600;
        }}
        .header p {{
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 16px;
        }}
        .stats {{
            display: flex;
            justify-content: space-around;
            padding: 20px;
            background-color: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
        }}
        .stat-box {{
            text-align: center;
            padding: 10px;
        }}
        .stat-number {{
            font-size: 32px;
            font-weight: bold;
            color: #667eea;
        }}
        .stat-label {{
            font-size: 14px;
            color: #6c757d;
            margin-top: 5px;
        }}
        .section {{
            padding: 30px;
        }}
        .section-title {{
            font-size: 22px;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
            display: flex;
            align-items: center;
        }}
        .section-title.vip {{
            border-bottom-color: #dc3545;
            color: #dc3545;
        }}
        .section-title .icon {{
            margin-right: 10px;
            font-size: 24px;
        }}
        .email-card {{
            background-color: #fff;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 15px;
            transition: box-shadow 0.3s;
        }}
        .email-card:hover {{
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }}
        .email-card.vip {{
            border-left: 4px solid #dc3545;
            background-color: #fff5f5;
        }}
        .email-header {{
            display: flex;
            justify-content: space-between;
            align-items: start;
            margin-bottom: 10px;
        }}
        .sender-info {{
            flex: 1;
        }}
        .sender-name {{
            font-weight: 600;
            font-size: 16px;
            color: #212529;
        }}
        .sender-email {{
            font-size: 13px;
            color: #6c757d;
            margin-top: 2px;
        }}
        .timestamp {{
            font-size: 13px;
            color: #6c757d;
            white-space: nowrap;
            margin-left: 15px;
        }}
        .subject {{
            font-size: 15px;
            font-weight: 600;
            color: #495057;
            margin: 10px 0;
        }}
        .preview {{
            font-size: 14px;
            color: #6c757d;
            line-height: 1.5;
            margin-top: 10px;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 5px;
        }}
        .badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
            margin-left: 10px;
        }}
        .badge.unread {{
            background-color: #0d6efd;
            color: white;
        }}
        .sender-group {{
            margin-bottom: 25px;
        }}
        .sender-group-header {{
            font-weight: 600;
            font-size: 15px;
            color: #495057;
            margin-bottom: 10px;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 5px;
        }}
        .email-list-item {{
            padding: 10px 15px;
            border-left: 3px solid #667eea;
            background-color: #f8f9fa;
            margin-bottom: 8px;
            border-radius: 4px;
        }}
        .email-list-item .subject {{
            margin: 0;
            font-size: 14px;
        }}
        .email-list-item .timestamp {{
            margin-left: 0;
            margin-top: 5px;
            display: block;
        }}
        .footer {{
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #6c757d;
            font-size: 13px;
            border-top: 2px solid #e9ecef;
        }}
        .empty-state {{
            text-align: center;
            padding: 40px;
            color: #6c757d;
        }}
        .empty-state .icon {{
            font-size: 48px;
            margin-bottom: 10px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📧 Resumen Diario de Correos</h1>
            <p>{date_range_str}</p>
        </div>
        
        <div class="stats">
            <div class="stat-box">
                <div class="stat-number">{len(equipo_directo)}</div>
                <div class="stat-label">Equipo Directo</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{len(otros_prioritarios)}</div>
                <div class="stat-label">Otros Prioritarios</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{filtered_count}</div>
                <div class="stat-label">🚫 Filtrados</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{total_conversations}</div>
                <div class="stat-label">Conversaciones</div>
            </div>
        </div>
"""
        
        # Sección Equipo Directo (VIP)
        html_output += """
        <div class="section">
            <div class="section-title vip">
                <span class="icon">⭐</span>
                <span>Equipo Directo (VIP)</span>
            </div>
"""
        
        if equipo_directo:
            for email in equipo_directo:
                unread_badge = '<span class="badge unread">No leído</span>' if email['is_unread'] else ''
                time_str = email['received_time'].strftime('%H:%M')
                
                # Escapar caracteres HTML para evitar inyección y rotura de estructura
                sender_name = html.escape(email['sender_name'])
                sender_email = html.escape(email['sender_email'])
                subject = html.escape(email['subject'])
                
                # Limpiar el preview del body y escapar HTML
                preview = html.escape(email['body']).replace('\r\n', ' ').replace('\n', ' ').strip()
                if len(preview) > 200:
                    preview = preview[:200] + '...'
                
                # NO indentar el HTML para evitar anidación visual
                html_output += f"""
<div class="email-card vip">
    <div class="email-header">
        <div class="sender-info">
            <div class="sender-name">{sender_name}{unread_badge}</div>
            <div class="sender-email">{sender_email}</div>
        </div>
        <div class="timestamp">🕐 {time_str}</div>
    </div>
    <div class="subject">{subject}</div>
    <div class="preview">{preview}</div>
</div>
"""
        else:
            html_output += """
            <div class="empty-state">
                <div class="icon">✨</div>
                <p>No hay correos del equipo directo hoy</p>
            </div>
"""
        
        html_output += """
        </div>
"""
        
        # Sección Otros Prioritarios
        html_output += """
        <div class="section">
            <div class="section-title">
                <span class="icon">📬</span>
                <span>Otros Prioritarios</span>
            </div>
"""
        
        if otros_prioritarios:
            for sender_name, emails in sorted(otros_by_sender.items()):
                email_count = len(emails)
                # Escapar el nombre del remitente
                sender_name_escaped = html.escape(sender_name)
                html_output += f"""
            <div class="sender-group">
                <div class="sender-group-header">
                    {sender_name_escaped} ({email_count} correo{'s' if email_count > 1 else ''})
                </div>
"""
                for email in emails:
                    unread_badge = '<span class="badge unread">No leído</span>' if email['is_unread'] else ''
                    time_str = email['received_time'].strftime('%H:%M')
                    
                    # Escapar subject y sender_email
                    subject_escaped = html.escape(email['subject'])
                    sender_email_escaped = html.escape(email['sender_email'])
                    
                    # Limpiar el preview del body y escapar HTML (igual que en VIP)
                    preview = html.escape(email['body']).replace('\r\n', ' ').replace('\n', ' ').strip()
                    if len(preview) > 200:
                        preview = preview[:200] + '...'
                    
                    html_output += f"""
                <div class="email-list-item">
                    <div class="subject">{subject_escaped}{unread_badge}</div>
                    <div class="timestamp">🕐 {time_str} • {sender_email_escaped}</div>
                    <div class="preview">{preview}</div>
                </div>
"""
                
                html_output += """
            </div>
"""
        else:
            html_output += """
            <div class="empty-state">
                <div class="icon">✨</div>
                <p>No hay otros correos prioritarios hoy</p>
            </div>
"""
        
        html_output += """
        </div>
"""
        
        # Footer
        html_output += f"""
        <div class="footer">
            <p>Este resumen fue generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}</p>
            <p>Los correos mostrados son únicos por conversación (se muestra solo el más reciente de cada hilo)</p>
        </div>
    </div>
</body>
</html>
"""
        
        return html_output
    
    def send_digest_email(self, html_content, subject=None):
        """
        Envía el resumen por correo electrónico.
        
        Args:
            html_content: Contenido HTML del correo
            subject: Asunto del correo (opcional)
        """
        if subject is None:
            today = datetime.now().strftime('%d/%m/%Y')
            subject = f"📧 Resumen Diario - {today}"
        
        try:
            # Crear nuevo correo
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = self.GMAIL_RECIPIENT
            mail.Subject = subject
            mail.HTMLBody = html_content
            
            # Enviar
            mail.Send()
            
            print(f"\n✓ Resumen enviado exitosamente a {self.GMAIL_RECIPIENT}")
            return True
            
        except Exception as e:
            print(f"\n✗ Error al enviar el correo: {e}")
            return False
    
    def generate_and_send_digest(self, delete_otros=False, start_date=None, end_date=None):
        """
        Genera y envía el resumen diario o de un rango de fechas.
        
        Args:
            delete_otros: Si True, elimina los correos clasificados como "Otros"
            start_date: Fecha inicial (datetime.date). Si es None, usa hoy.
            end_date: Fecha final (datetime.date). Si es None, usa hoy.
        """
        print("\n" + "=" * 70)
        print("📧 GENERADOR DE RESUMEN DIARIO - OUTLOOK")
        print("=" * 70)
        print(f"⏰ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # Obtener correos del rango de fechas
        email_data = self.get_todays_emails(delete_otros=delete_otros, start_date=start_date, end_date=end_date)
        
        # Determinar string de rango de fechas para el reporte
        if start_date is None:
            start_date = datetime.now().date()
        if end_date is None:
            end_date = datetime.now().date()
        
        if start_date == end_date:
            date_range_str = start_date.strftime('%d/%m/%Y')
            subject = f"📧 Resumen Diario - {date_range_str}"
        else:
            date_range_str = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
            subject = f"📧 Resumen de Correos - {date_range_str}"
        
        # Generar HTML
        print("\n🎨 Generando reporte HTML...")
        html_content = self.generate_html_report(email_data, date_range_str=date_range_str)
        
        # Enviar correo
        print(f"📤 Enviando resumen a {self.GMAIL_RECIPIENT}...")
        success = self.send_digest_email(html_content, subject=subject)
        
        if success:
            print("\n" + "=" * 70)
            print("✅ RESUMEN DIARIO COMPLETADO")
            print("=" * 70)
        else:
            print("\n" + "=" * 70)
            print("❌ ERROR AL ENVIAR RESUMEN")
            print("=" * 70)
        
        return success


def main():
    """Función principal."""
    import argparse
    
    # Configurar argumentos de línea de comandos
    parser = argparse.ArgumentParser(
        description='Generador de Resumen Diario de Correos de Outlook',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python outlook_daily_digest.py                                    # Modo interactivo
  python outlook_daily_digest.py --delete                           # Hoy con borrado
  python outlook_daily_digest.py --no-delete                        # Hoy sin borrado
  python outlook_daily_digest.py --date 2026-01-14 --delete         # Día específico
  python outlook_daily_digest.py --start-date 2026-01-14 --end-date 2026-01-15 --delete  # Rango
        """
    )
    
    parser.add_argument(
        '--delete',
        action='store_true',
        help='Eliminar correos antiguos de conversaciones y filtrados (modo automático)'
    )
    
    parser.add_argument(
        '--no-delete',
        action='store_true',
        help='NO eliminar correos, solo generar resumen (modo automático)'
    )
    
    parser.add_argument(
        '--date',
        type=str,
        help='Procesar un día específico (formato: YYYY-MM-DD)'
    )
    
    parser.add_argument(
        '--start-date',
        type=str,
        help='Fecha inicial del rango (formato: YYYY-MM-DD)'
    )
    
    parser.add_argument(
        '--end-date',
        type=str,
        help='Fecha final del rango (formato: YYYY-MM-DD)'
    )
    
    args = parser.parse_args()
    
    # Parsear fechas si se proporcionaron
    start_date = None
    end_date = None
    
    if args.date:
        # Si se especifica --date, usar esa fecha para start y end
        try:
            start_date = datetime.strptime(args.date, '%Y-%m-%d').date()
            end_date = start_date
        except ValueError:
            print(f"❌ Error: Formato de fecha inválido '{args.date}'. Use YYYY-MM-DD")
            sys.exit(1)
    elif args.start_date or args.end_date:
        # Si se especifica rango, parsear ambas fechas
        if args.start_date:
            try:
                start_date = datetime.strptime(args.start_date, '%Y-%m-%d').date()
            except ValueError:
                print(f"❌ Error: Formato de fecha inválido '{args.start_date}'. Use YYYY-MM-DD")
                sys.exit(1)
        
        if args.end_date:
            try:
                end_date = datetime.strptime(args.end_date, '%Y-%m-%d').date()
            except ValueError:
                print(f"❌ Error: Formato de fecha inválido '{args.end_date}'. Use YYYY-MM-DD")
                sys.exit(1)
        
        # Si solo se especificó una fecha, usar la misma para ambas
        if start_date and not end_date:
            end_date = start_date
        elif end_date and not start_date:
            start_date = end_date
    
    # Determinar modo de operación
    if args.delete and args.no_delete:
        print("❌ Error: No puedes usar --delete y --no-delete al mismo tiempo")
        sys.exit(1)
    
    # Modo automático (con argumentos)
    if args.delete or args.no_delete:
        delete_otros = args.delete
        print("\n" + "=" * 70)
        print("📧 RESUMEN DIARIO DE CORREOS - MODO AUTOMÁTICO")
        print("=" * 70)
        print(f"Modo: {'BORRADO ACTIVADO' if delete_otros else 'SOLO RESUMEN'}\n")
    
    # Modo interactivo (sin argumentos)
    else:
        print("\n" + "=" * 70)
        print("📧 RESUMEN DIARIO DE CORREOS")
        print("=" * 70)
        print("\nEste script generará un resumen de los correos recibidos hoy")
        print("y lo enviará a tu Gmail.\n")
        
        # Preguntar si eliminar "Otros"
        print("¿Deseas eliminar los correos clasificados como 'Otros' (no prioritarios)?")
        print("  1. No, solo generar resumen")
        print("  2. Sí, eliminar correos 'Otros' después de generar resumen")
        
        while True:
            choice = input("\nSelecciona opción (1 o 2): ").strip()
            if choice == "1":
                delete_otros = False
                break
            elif choice == "2":
                delete_otros = True
                print("\n⚠️  ADVERTENCIA: Los correos 'Otros' serán ELIMINADOS")
                confirm = input("¿Estás seguro? (s/n): ").strip().lower()
                if confirm == 's':
                    break
                else:
                    print("Operación cancelada. Solo se generará el resumen.\n")
                    delete_otros = False
                    break
            else:
                print("Por favor selecciona 1 o 2.")
    
    # Crear generador de resumen
    digest = OutlookDailyDigest()
    
    # Generar y enviar
    digest.generate_and_send_digest(delete_otros=delete_otros, start_date=start_date, end_date=end_date)


if __name__ == "__main__":
    main()

