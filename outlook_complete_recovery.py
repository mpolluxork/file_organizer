"""
Script de recuperación completa - Lee correos del Inbox Y de Elementos eliminados.
Este script NO borra ni marca nada, solo genera el resumen completo.
"""

import win32com.client
from datetime import datetime, timedelta
from collections import defaultdict
import sys
import json
import os
import html

# Importar la clase original
import outlook_daily_digest

class CompleteRecoveryDigest(outlook_daily_digest.OutlookDailyDigest):
    """Versión que lee del Inbox Y de Elementos eliminados."""
    
    def get_deleted_items_folder(self):
        """Obtiene la carpeta de Elementos eliminados."""
        try:
            return self.namespace.GetDefaultFolder(3)  # 3 = olFolderDeletedItems
        except Exception as e:
            print(f"✗ Error al acceder a Elementos eliminados: {e}")
            return None
    
    def get_todays_emails(self, delete_otros=False, start_date=None, end_date=None):
        """
        Lee correos del Inbox Y de Elementos eliminados.
        """
        # Forzar delete_otros a False
        delete_otros = False
        
        # Determinar rango de fechas
        if start_date is None:
            start_date = datetime.now().date()
        if end_date is None:
            end_date = datetime.now().date()
        
        if start_date > end_date:
            print(f"❌ Error: La fecha inicial ({start_date}) no puede ser posterior a la fecha final ({end_date})")
            sys.exit(1)
        
        filter_start = start_date - timedelta(days=1)
        filter_end = end_date + timedelta(days=1)
        filter_string = f"[ReceivedTime] > '{filter_start.strftime('%d/%m/%Y')}' AND [ReceivedTime] < '{filter_end.strftime('%d/%m/%Y')}'"
        
        print(f"\n📊 MODO RECUPERACIÓN COMPLETA")
        if start_date == end_date:
            print(f"   Fecha: {start_date.strftime('%Y-%m-%d')}")
        else:
            print(f"   Rango: {start_date.strftime('%Y-%m-%d')} al {end_date.strftime('%Y-%m-%d')}")
        print(f"   ⚠️  ESTE SCRIPT NO BORRARÁ NI MARCARÁ NADA\n")
        
        # Leer del Inbox
        inbox = self.get_inbox()
        inbox_messages = inbox.Items.Restrict(filter_string)
        print(f"   📥 Inbox: {inbox_messages.Count} correos")
        
        # Leer de Elementos eliminados
        deleted_folder = self.get_deleted_items_folder()
        deleted_count = 0
        deleted_messages = []
        if deleted_folder:
            deleted_messages = deleted_folder.Items.Restrict(filter_string)
            deleted_count = deleted_messages.Count
            print(f"   🗑️  Elementos eliminados: {deleted_count} correos")
        
        total_found = inbox_messages.Count + deleted_count
        print(f"   📊 TOTAL: {total_found} correos\n")
        
        # Combinar ambas listas
        all_messages = list(inbox_messages) + list(deleted_messages)
        
        # Procesar todos los mensajes
        conversations = {}
        total_processed = 0
        filtered_count = 0
        
        for message in all_messages:
            try:
                total_processed += 1
                
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
                
                # Aplicar filtros personalizados
                if self.should_filter_email(sender_email, subject):
                    filtered_count += 1
                    continue
                
                email_data = {
                    'sender_email': sender_email,
                    'sender_name': sender_name,
                    'subject': subject,
                    'received_time': received_time,
                    'body': body[:500],
                    'is_unread': is_unread,
                    'conversation_id': conversation_id,
                    'has_attachments': message.Attachments.Count > 0,
                    'message_object': message
                }
                
                # Agrupar por conversación
                if conversation_id:
                    if conversation_id not in conversations:
                        conversations[conversation_id] = {
                            'latest_message': email_data,
                            'has_vip': sender_email in self.EQUIPO_DIRECTO,
                            'all_messages': [email_data]
                        }
                    else:
                        conversations[conversation_id]['all_messages'].append(email_data)
                        
                        if received_time > conversations[conversation_id]['latest_message']['received_time']:
                            conversations[conversation_id]['latest_message'] = email_data
                        
                        if sender_email in self.EQUIPO_DIRECTO:
                            conversations[conversation_id]['has_vip'] = True
                else:
                    unique_key = f"no_conv_{total_processed}"
                    conversations[unique_key] = {
                        'latest_message': email_data,
                        'has_vip': sender_email in self.EQUIPO_DIRECTO,
                        'all_messages': [email_data]
                    }
                
            except Exception as e:
                print(f"  ⚠ Error procesando correo: {e}")
                continue
        
        # Procesar conversaciones
        final_emails = []
        
        for conv_id, conv_data in conversations.items():
            latest_message = conv_data['latest_message']
            latest_message['is_vip_conversation'] = conv_data['has_vip']
            final_emails.append(latest_message)
        
        # Clasificar
        equipo_directo_emails = []
        otros_prioritarios = []
        
        for email_data in final_emails:
            if email_data.get('is_vip_conversation', False):
                equipo_directo_emails.append(email_data)
            else:
                otros_prioritarios.append(email_data)
        
        equipo_directo_emails.sort(key=lambda x: x['received_time'], reverse=True)
        otros_prioritarios.sort(key=lambda x: x['received_time'], reverse=True)
        
        print(f"   Conversaciones únicas: {len(conversations)}")
        print(f"   Equipo Directo: {len(equipo_directo_emails)}")
        print(f"   Otros Prioritarios: {len(otros_prioritarios)}")
        print(f"   🚫 Filtrados: {filtered_count}")
        
        return {
            'equipo_directo': equipo_directo_emails,
            'otros_prioritarios': otros_prioritarios,
            'filtered_count': filtered_count,
            'deleted_count': 0,
            'conversation_deleted_count': 0,
            'total_conversations': len(conversations),
            'total_processed': total_processed
        }


def main():
    """Función principal."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Recuperación Completa - Lee Inbox + Elementos eliminados',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplo:
  python outlook_complete_recovery.py --start-date 2026-01-14 --end-date 2026-01-15
        """
    )
    
    parser.add_argument('--date', type=str, help='Día específico (YYYY-MM-DD)')
    parser.add_argument('--start-date', type=str, help='Fecha inicial (YYYY-MM-DD)')
    parser.add_argument('--end-date', type=str, help='Fecha final (YYYY-MM-DD)')
    
    args = parser.parse_args()
    
    start_date = None
    end_date = None
    
    if args.date:
        try:
            start_date = datetime.strptime(args.date, '%Y-%m-%d').date()
            end_date = start_date
        except ValueError:
            print(f"❌ Error: Formato inválido '{args.date}'. Use YYYY-MM-DD")
            sys.exit(1)
    elif args.start_date or args.end_date:
        if args.start_date:
            try:
                start_date = datetime.strptime(args.start_date, '%Y-%m-%d').date()
            except ValueError:
                print(f"❌ Error: Formato inválido '{args.start_date}'. Use YYYY-MM-DD")
                sys.exit(1)
        
        if args.end_date:
            try:
                end_date = datetime.strptime(args.end_date, '%Y-%m-%d').date()
            except ValueError:
                print(f"❌ Error: Formato inválido '{args.end_date}'. Use YYYY-MM-DD")
                sys.exit(1)
        
        if start_date and not end_date:
            end_date = start_date
        elif end_date and not start_date:
            start_date = end_date
    else:
        print("❌ Error: Especifica --date o --start-date/--end-date")
        sys.exit(1)
    
    print("\n" + "=" * 70)
    print("🔄 RECUPERACIÓN COMPLETA - INBOX + ELEMENTOS ELIMINADOS")
    print("=" * 70)
    print("⚠️  Este script NO borrará ni marcará ningún correo")
    print("=" * 70)
    
    digest = CompleteRecoveryDigest()
    digest.generate_and_send_digest(delete_otros=False, start_date=start_date, end_date=end_date)


if __name__ == "__main__":
    main()
