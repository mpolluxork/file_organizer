"""
Script temporal para recuperar TODOS los correos (leídos y no leídos) de un rango de fechas.
Este script NO borra ni marca nada, solo genera el resumen.
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

class RecoveryDigest(outlook_daily_digest.OutlookDailyDigest):
    """Versión de recuperación que procesa TODOS los correos, leídos o no."""
    
    def get_todays_emails(self, delete_otros=False, start_date=None, end_date=None):
        """
        Versión modificada que NO filtra por estado de lectura y NO borra nada.
        """
        # Forzar delete_otros a False para no borrar nada
        delete_otros = False
        
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
        filter_start = start_date - timedelta(days=1)
        filter_end = end_date + timedelta(days=1)
        
        filter_string = f"[ReceivedTime] > '{filter_start.strftime('%d/%m/%Y')}' AND [ReceivedTime] < '{filter_end.strftime('%d/%m/%Y')}'"
        messages = messages.Restrict(filter_string)
        messages.Sort("[ReceivedTime]", True)
        
        print(f"\n📊 MODO RECUPERACIÓN: Analizando TODOS los correos (leídos y no leídos)")
        if start_date == end_date:
            print(f"   Fecha: {start_date.strftime('%Y-%m-%d')}")
        else:
            print(f"   Rango: {start_date.strftime('%Y-%m-%d')} al {end_date.strftime('%Y-%m-%d')}")
        print(f"   Total de correos encontrados: {messages.Count}")
        print(f"   ⚠️  ESTE SCRIPT NO BORRARÁ NI MARCARÁ NADA\n")
        
        # Categorizar correos
        equipo_directo_emails = []
        otros_prioritarios = []
        filtered_count = 0
        
        conversations = {}
        total_processed = 0
        messages_list = list(messages)
        
        for message in messages_list:
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
                
                # Aplicar filtros personalizados (UTEL, LinkedIn, asuntos, etc.)
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
        
        # Clasificar correos finales
        equipo_directo_emails = []
        otros_prioritarios = []
        
        for email_data in final_emails:
            if email_data.get('is_vip_conversation', False):
                equipo_directo_emails.append(email_data)
            else:
                otros_prioritarios.append(email_data)
        
        # Ordenar por fecha
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
    """Función principal para recuperación."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Script de Recuperación - Genera resumen de TODOS los correos (leídos y no leídos)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python outlook_recovery.py --start-date 2026-01-14 --end-date 2026-01-15
  python outlook_recovery.py --date 2026-01-14
        """
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
    
    # Parsear fechas
    start_date = None
    end_date = None
    
    if args.date:
        try:
            start_date = datetime.strptime(args.date, '%Y-%m-%d').date()
            end_date = start_date
        except ValueError:
            print(f"❌ Error: Formato de fecha inválido '{args.date}'. Use YYYY-MM-DD")
            sys.exit(1)
    elif args.start_date or args.end_date:
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
        
        if start_date and not end_date:
            end_date = start_date
        elif end_date and not start_date:
            start_date = end_date
    else:
        print("❌ Error: Debes especificar --date o --start-date/--end-date")
        sys.exit(1)
    
    print("\n" + "=" * 70)
    print("🔄 SCRIPT DE RECUPERACIÓN - MODO SOLO LECTURA")
    print("=" * 70)
    print("⚠️  Este script NO borrará ni marcará ningún correo")
    print("=" * 70)
    
    # Crear generador de resumen
    digest = RecoveryDigest()
    
    # Generar y enviar (sin borrar nada)
    digest.generate_and_send_digest(delete_otros=False, start_date=start_date, end_date=end_date)


if __name__ == "__main__":
    main()
