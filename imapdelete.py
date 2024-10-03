#!/usr/bin/env python3.11

"""
IMAP Email Manager

Questo programma permette di gestire le email su un server IMAP, offrendo funzionalità
per elencare le cartelle, cercare e cancellare messaggi in base a criteri specifici.

Funzionalità principali:
- Elencare le cartelle IMAP disponibili
- Cercare messaggi in una cartella specifica
- Filtrare i messaggi per data e oggetto (usando espressioni regolari)
- Cancellare o spostare nel cestino i messaggi selezionati
- Visualizzare un'anteprima dei messaggi prima della cancellazione
- Mostrare una barra di avanzamento durante la ricerca e la cancellazione
- Archiviare i messaggi in una struttura di cartelle IMAP
- Archiviare i messaggi in una struttura di cartelle locali
- Modalità di debug per la risoluzione dei problemi

Autore: Diego Ercolani
Data: 2/10/2024 ultima modifica 2/10/2024
Versione: 1.0
Licenza: eredita le licenze delle librerie utilizzate, la mia parte è frutto di
         collaborazioni multiple quindi GPL
"""

import argparse
import imaplib
import re
import getpass
import sys
from datetime import datetime
import shutil
import email
from email.header import decode_header
from email import message_from_bytes
import os
import time

def print_help():
    """
    Stampa la pagina di aiuto interattiva che descrive i casi d'uso del programma.
    """
    print("""
IMAP Email Manager - Guida Interattiva

Questo programma ti permette di gestire le tue email su un server IMAP. Ecco i principali casi d'uso:

1. Elencare le cartelle IMAP disponibili:
   python imap_manager.py -u username@example.com -s imap.example.com -l

2. Cercare e cancellare messaggi in una cartella specifica:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -d "01/01/2023-31/12/2023" "oggetto da cercare"

3. Spostare i messaggi nel cestino invece di cancellarli definitivamente:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -d "01/01/2023-31/12/2023" "oggetto da cercare"

4. Cancellare definitivamente i messaggi (usa con cautela):
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -d "01/01/2023-31/12/2023" -e "oggetto da cercare"
   
5. Cercare messaggi con criteri specifici negli header:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -a "From" "example\.com$" -o "X-Spam-Flag" "YES"
   
6. Archiviare i messaggi in una struttura di cartelle IMAP:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" --archive "Archive" "oggetto da cercare"

7. Archiviare i messaggi in una struttura di cartelle locali:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" --archive-to-disk "/path/to/archive" "oggetto da cercare"

Parametri:
-u: Username per l'accesso IMAP
-s: Server IMAP (con porta opzionale, es. imap.example.com:993)
-p: Password (se omessa, verrà richiesta in modo sicuro)
-l: Elenca le cartelle IMAP disponibili
-f: Nome della cartella IMAP da utilizzare
-d: Intervallo di date per la ricerca (formato: "gg/mm/aaaa-gg/mm/aaaa")
-e: Cancella definitivamente i messaggi invece di spostarli nel cestino
[regex]: Espressione regolare opzionale per filtrare i messaggi per oggetto
-a, --and-header: Ricerca AND nell'header specificato usando la regex fornita (può essere usato più volte)
-o, --or-header: Ricerca OR nell'header specificato usando la regex fornita (può essere usato più volte)

--archive: Specifica la cartella IMAP di destinazione per l'archiviazione dei messaggi
--archive-to-disk: Specifica la cartella locale di destinazione per l'archiviazione dei messaggi
--debug: Abilita i messaggi di debug

Per ulteriori informazioni su un parametro specifico, digita il nome del parametro (es. '-u'):
""")
    
    while True:
        user_input = input("> ").strip().lower()
        if user_input == 'q' or user_input == 'quit' or user_input == 'exit':
            break
        elif user_input == '-u':
            print("-u: Specifica l'username per l'accesso al server IMAP.")
        elif user_input == '-s':
            print("-s: Indica l'indirizzo del server IMAP, opzionalmente seguito dalla porta (es. imap.example.com:993).")
        elif user_input == '-p':
            print("-p: Password per l'accesso. Se omessa, verrà richiesta in modo sicuro durante l'esecuzione.")
        elif user_input == '-l':
            print("-l: Elenca tutte le cartelle IMAP disponibili nell'account.")
        elif user_input == '-f':
            print("-f: Specifica il nome della cartella IMAP su cui operare.")
        elif user_input == '-d':
            print("-d: Imposta l'intervallo di date per la ricerca dei messaggi (formato: 'gg/mm/aaaa-gg/mm/aaaa').")
        elif user_input == '-e':
            print("-e: Se presente, i messaggi verranno cancellati definitivamente invece di essere spostati nel cestino.")
        elif user_input == '--archive':
            print("--archive: Specifica la cartella IMAP di destinazione per l'archiviazione dei messaggi. "
                  "I messaggi verranno archiviati in una struttura di cartelle organizzata per anno e mese. "
                  "Con questa opzione, i messaggi vengono spostati e non copiati nel cestino.")
        elif user_input == '--archive-to-disk':
            print("--archive-to-disk: Specifica la cartella locale di destinazione per l'archiviazione dei messaggi. "
                  "I messaggi verranno archiviati in una struttura di cartelle locale organizzata per anno e mese. "
                  "Con questa opzione, i messaggi vengono archiviati localmente e spostati nel cestino, a meno che non sia presente il flag -e.")
        elif user_input == '--debug':
            print("--debug: "
                  "Attiva modalità debug. "
                  )
        else:
            print("Parametro non riconosciuto. Prova con -u, -s, -p, -l, -f, -d, -e, --archive, o --archive-to-disk.")
        print("\nInserisci un altro parametro o 'q' per uscire:")

def create_imap_folder(imap, folder_name, user, debug=False):
    if debug:
        print(f"DEBUG: Tentativo di creare la cartella: {folder_name}")
    try:
        # Rimuovi le virgolette dal nome della cartella
        folder_name = folder_name.strip('"')
        if debug:
            print(f"DEBUG: Nome cartella senza virgolette: {folder_name}")

        # Dividi il percorso della cartella in componenti
        folder_parts = folder_name.split('/')
        if debug:
            print(f"DEBUG: Componenti del percorso: {folder_parts}")

        # Crea ogni livello della cartella se non esiste
        current_path = ''
        for part in folder_parts:
            if current_path:
                current_path += '/'
            current_path += part
            if debug:
                print(f"DEBUG: Verifica/creazione del percorso: {current_path}")

            # Verifica se la cartella esiste
            res, folders = imap.list(directory=f'"{current_path}"',pattern='*')
            if debug:
                print(f"DEBUG: Risultato list per {current_path}: res={res}")

            # Gestisci il caso in cui 'folders' è None o [None]
            if res != 'OK' or not folders or folders == [None] :
                if debug:
                    print(f"DEBUG: La cartella {current_path} non esiste, tentativo di creazione...")
                # Se la cartella non esiste, la creiamo
                res, data = imap.create(current_path)
                if debug:
                    print(f"DEBUG: Risultato create per {current_path}: res={res}, data={data}")
                if res != 'OK':
                    print(f"Errore nella creazione della cartella {current_path}")
                    return False
            else:
                if debug:
                    print(f"DEBUG: La cartella {current_path} esiste già")
                    print(f"DEBUG: {folders} {type(folders)}")
                # Verifica se la cartella ha il flag \\HasChildren
                # Filtra eventuali valori None in 'folders'
                valid_folders = [folder for folder in folders if folder is not None]
                if not any(b'\\HasChildren' in folder for folder in valid_folders):
                    if debug:
                        print(f"DEBUG: La cartella {current_path} non ha il flag \\HasChildren, aggiornamento...")
                    # Aggiorna i flag della cartella (se necessario)
                    # Nota: potrebbe essere necessario un comando specifico per aggiornare i flag
    except imaplib.IMAP4.error as e:
        print(f"Errore IMAP durante la creazione della cartella {folder_name}: {str(e)}")
        return False
    if debug:
        print(f"DEBUG: Creazione della cartella {folder_name} completata con successo")
    return True


def get_hierarchy_delimiter(imap):
    result, data = imap.list()
    if result == 'OK':
        # Esempio di risposta: '("*" "." "")'
        # Estrai il delimitatore di gerarchia dalla risposta
        hierarchy_delimiter = data[0].decode().split(' ')[2].strip('"')
        print(f"Delimitatore di gerarchia: {hierarchy_delimiter}")
    else:
        print("Impossibile ottenere il delimitatore di gerarchia dal server IMAP.")
        hierarchy_delimiter = '/'
    return hierarchy_delimiter


def archive_message_imap(imap, msg_id, dest_folder, source_folder, user, debug=False):
    if debug:
        print(f"DEBUG: Inizio archiviazione del messaggio {msg_id}")
    res, msg_data = imap.fetch(msg_id, '(RFC822)')
    if res != 'OK':
        print(f"Errore nel recupero del messaggio {msg_id}")
        return False

    email_body = msg_data[0][1]
    email_message = message_from_bytes(email_body)
    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    
    source_folder_name = os.path.basename(source_folder.strip('"'))
    if debug:
        print(f"DEBUG: Nome cartella sorgente: {source_folder_name}")
    if date_tuple:
        year = str(date_tuple[0])
        month = f"{date_tuple[1]:02d}"
        archive_path = f'{dest_folder}/{source_folder_name}/{year}/{month}'
    else:
        archive_path = f'{dest_folder}/{source_folder_name}'
    
    if debug:
        print(f"DEBUG: Percorso di archiviazione: {archive_path}")
    if not create_imap_folder(imap, archive_path, user, debug):
        if debug:
            print(f"DEBUG: Fallimento nella creazione della cartella {archive_path}")
        return False
    
    # Verifica se la cartella esiste prima di tentare l'append
    res, _ = imap.list(f'"{archive_path}"')
    if res != 'OK':
        print(f"Errore: La cartella {archive_path} non esiste o non è accessibile")
        return False
    
    if debug:
        print(f"DEBUG: Tentativo di append del messaggio in {archive_path}")
    #res = imap.append(f'"{archive_path}"', '', imaplib.Time2Internaldate(time.time()), email_body)
    res = imap.append(archive_path, '', imaplib.Time2Internaldate(time.time()), email_body)

    if debug:
        print(f"DEBUG: Risultato append: {res}")
    return res[0] == 'OK'



def archive_message_disk(msg_id, imap, dest_folder, source_folder, debug=False):
    if debug:
        print(f"DEBUG: Inizio archiviazione su disco del messaggio {msg_id}")
    
    res, msg_data = imap.fetch(msg_id, '(RFC822)')
    if res != 'OK':
        print(f"Errore nel recupero del messaggio {msg_id}")
        return False

    email_body = msg_data[0][1]
    email_message = email.message_from_bytes(email_body)
    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    
    if debug:
        print(f"DEBUG: Data del messaggio: {email_message['Date']}")
    
    source_folder_name = os.path.basename(source_folder.strip('"'))
    if date_tuple:
        year = str(date_tuple[0])
        month = f"{date_tuple[1]:02d}"
        archive_path = os.path.join(dest_folder, source_folder_name, year, month)
    else:
        archive_path = os.path.join(dest_folder, source_folder_name)
    
    if debug:
        print(f"DEBUG: Percorso di archiviazione: {archive_path}")
    
    os.makedirs(archive_path, exist_ok=True)
    
    subject = email_message['Subject']
    if subject:
        subject = decode_mime_words(subject)
    else:
        subject = 'No Subject'
    
    safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)
    safe_filename = f"{email_message['Date']}_{safe_subject[:50]}.eml"
    safe_filename = safe_filename.replace(':', '_')
    
    file_path = os.path.join(archive_path, safe_filename)
    
    if debug:
        print(f"DEBUG: Salvataggio del messaggio in: {file_path}")
    
    try:
        with open(file_path, 'wb') as f:
            f.write(email_body)
        if debug:
            print(f"DEBUG: Messaggio salvato con successo")
        return True
    except IOError as e:
        print(f"Errore durante il salvataggio del messaggio: {e}")
        if debug:
            print(f"DEBUG: Errore dettagliato: {str(e)}")
        return False








def decode_mime_words(s):
    return ''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in decode_header(s)
    )

def show_grouped_subjects_and_select(filtered_msgs):
    subject_count = {}
    for msg_id, subject, date in filtered_msgs:
        if subject not in subject_count:
            subject_count[subject] = {'count': 0, 'ids': [], 'dates': []}
        subject_count[subject]['count'] += 1
        subject_count[subject]['ids'].append(msg_id)
        subject_count[subject]['dates'].append(date)
    
    print(f"\nSoggetti dei messaggi trovati (totale: {len(filtered_msgs)}):")
    subjects_list = sorted(subject_count.items(), key=lambda x: x[1]['count'], reverse=True)
    for i, (subject, data) in enumerate(subjects_list, 1):
        dates = sorted(data['dates'])
        if len(dates) > 1:
            date_info = f"dal {dates[0]} al {dates[-1]}"
        else:
            date_info = f"il {dates[0]}"
        print(f"{i}. {subject} ({data['count']} messaggi) - {date_info}")
    
    selected_groups = []
    while True:
        choice = input("\nInserisci i numeri dei gruppi da selezionare (separati da virgola), 'a' per tutti, o 'n' per nessuno: ").lower()
        if choice == 'n':
            break
        elif choice == 'a':
            selected_groups = list(range(1, len(subjects_list) + 1))
            break
        else:
            try:
                selected_groups = [int(x.strip()) for x in choice.split(',') if x.strip()]
                if all(1 <= x <= len(subjects_list) for x in selected_groups):
                    break
                else:
                    print("Alcuni numeri non sono validi. Riprova.")
            except ValueError:
                print("Input non valido. Inserisci numeri separati da virgola, 'a' o 'n'.")
    
    messages_to_delete = []
    for i in selected_groups:
        subject, data = subjects_list[i-1]
        messages_to_delete.extend(data['ids'])
        print(f"Selezionato per la cancellazione: {subject} ({data['count']} messaggi)")
    
    return messages_to_delete




def create_progress_bar(total, current, matching, non_matching):
    width, _ = shutil.get_terminal_size()
    
    # Calcola lo spazio necessario per la percentuale e i caratteri accessori
    percent = f" {current/total*100:.1f}%"
    extra_chars = 2  # Per le parentesi quadre []
    
    # Sottrai lo spazio per la percentuale e i caratteri accessori
    available_width = width - len(percent) - extra_chars
    
    if available_width <= 0:
        return f"[{'*' * matching}{'_' * non_matching}{' ' * (total-current)}]{percent}"

    filled = int(available_width * current // total)
    matching_width = int(available_width * matching // total)
    non_matching_width = filled - matching_width
    remaining_width = available_width - filled

    matching_str = '*' * matching_width
    non_matching_str = '_' * non_matching_width
    remaining_str = ' ' * remaining_width

    # Aggiungi i contatori se c'è spazio sufficiente
    if matching > 0 and matching_width > len(str(matching)):
        matching_str = str(matching).center(matching_width, '*')
    if non_matching > 0 and non_matching_width > len(str(non_matching)):
        non_matching_str = str(non_matching).center(non_matching_width, '_')
    if (total-current) > 0 and remaining_width > len(str(total-current)):
        remaining_str = str(total-current).center(remaining_width, ' ')

    bar = matching_str + non_matching_str + remaining_str
    return f"[{bar}]{percent}"

def parse_args():
    parser = argparse.ArgumentParser(description='Script IMAP per gestione messaggi.')
    parser.add_argument('-l', '--list', action='store_true', help='Elenca le cartelle IMAP disponibili.')
    parser.add_argument('-f', '--folder', metavar='NOMECARTELLA', help='Nome della cartella IMAP da selezionare.')
    parser.add_argument('-s', '--server', metavar='SERVER[:PORTA]', required=True, help='Server IMAP e porta (opzionale).')
    parser.add_argument('-u', '--user', metavar='USERNAME', required=True, help='Username per l\'autenticazione.')
    parser.add_argument('-p', '--password', metavar='PASSWORD', nargs='?', help='Password per l\'autenticazione.')
    parser.add_argument('-d', '--datascope', metavar='INIZIO-FINE', help='Intervallo di date nel formato dd/mm/yyyy-dd/mm/yyyy.')
    parser.add_argument('-e', '--expunge', action='store_true', help='Cancella definitivamente i messaggi.')
    parser.add_argument('regex', nargs='?', help='Espressione regolare per filtrare i soggetti dei messaggi.')
    parser.add_argument('-a', '--and-header', action='append', nargs=2, metavar=('HEADER', 'REGEX'),
                        help='Ricerca AND nell\'header specificato usando la regex fornita')
    parser.add_argument('-o', '--or-header', action='append', nargs=2, metavar=('HEADER', 'REGEX'),
                        help='Ricerca OR nell\'header specificato usando la regex fornita')
    parser.add_argument('--archive', metavar='CARTELLA_DESTINAZIONE', help='Archivia spostandoli i messaggi nella cartella specificata')
    parser.add_argument('--archive-to-disk', metavar='CARTELLA_LOCALE_DESTINAZIONE', help='Archivia i messaggi nella cartella locale specificata')
    parser.add_argument('--debug', action='store_true', help='Abilita i messaggi di debug')
    return parser.parse_args()

def connect_imap(server, username, password):
    if ':' in server:
        server_name, port = server.split(':')
        port = int(port)
    else:
        server_name = server
        port = None
    try:
        if port:
            imap = imaplib.IMAP4_SSL(server_name, port)
        else:
            imap = imaplib.IMAP4_SSL(server_name)
        imap.login(username, password)
        return imap
    except imaplib.IMAP4.error as e:
        print(f'Errore durante la connessione al server IMAP: {e}')
        sys.exit(1)

def list_folders(imap):
    result, folders = imap.list()
    if result == 'OK':
        print('Cartelle disponibili:')
        for folder in folders:
            parts = folder.decode().split(' "/" ')
            if len(parts) == 2:
                print(parts[1].strip('"'))
            else:
                print(folder.decode())
    else:
        print('Impossibile recuperare le cartelle.')

def get_date_range(date_range_str):
    try:
        start_str, end_str = date_range_str.split('-')
        start_date = datetime.strptime(start_str.strip(), '%d/%m/%Y')
        end_date = datetime.strptime(end_str.strip(), '%d/%m/%Y')
        return start_date, end_date
    except ValueError:
        print('Formato delle date non valido. Utilizzare dd/mm/yyyy-dd/mm/yyyy.')
        sys.exit(1)
        
def get_header_value(header_data, header_name):
    match = re.search(f'{header_name}: (.*)', header_data, re.IGNORECASE | re.MULTILINE)
    if match:
        value = decode_mime_words(match.group(1).strip())
        if header_name.lower() == 'date':
            try:
                # Prova a parsare la data in un formato standard
                parsed_date = email.utils.parsedate_to_datetime(value)
                return parsed_date.strftime('%Y-%m-%d %H:%M:%S')
            except:
                # Se il parsing fallisce, restituisci la stringa originale
                return value
        return value
    return ''

def main():
    
    # Parse arguments
    args = parse_args()

    if not args.password:
        args.password = getpass.getpass('Inserisci la password: ')
        
    and_headers = args.and_header or []
    or_headers = args.or_header or []

    imap = connect_imap(args.server, args.user, args.password)

    if args.list:
        list_folders(imap)
        imap.logout()
        sys.exit(0)

    if not args.folder:
        print('Devi specificare una cartella con il parametro -f.')
        sys.exit(1)

    if args.debug:
        print("hierarchy_delimiter:"+get_hierarchy_delimiter(imap))
        result, data = imap.namespace()
        imap.debug = 4 # Livello di debug (0-5)
        if result == 'OK':
            print(f"Namespace: {data}")
            # Analizza il namespace per ottenere eventuali prefissi o delimitatori aggiuntivi
        else:
            print("Impossibile ottenere il namespace dal server IMAP.")

    folder_name = args.folder.replace('"', '') # gestisce caratteri speciali nel nome folder
    try:
        res, data = imap.select(folder_name)
    except imaplib.IMAP4.error as e:
        print(f'Errore nella selezione della cartella: {e}')
        print('Provo a utilizzare il nome della cartella tra virgolette...')
        try:
            res, data = imap.select(f'"{folder_name}"')
        except imaplib.IMAP4.error as e:
            print(f'Errore nella selezione della cartella: {e}')
            print('Impossibile selezionare la cartella. Verificare il nome e i permessi.')
            imap.logout()
            sys.exit(1)
            
    if res != 'OK':
        print(f'Impossibile selezionare la cartella "{folder_name}".')
        print(f'Errore: {data[0].decode()}')
        imap.logout()
        sys.exit(1)

    search_criteria = []

    if args.datascope:
        start_date, end_date = get_date_range(args.datascope)
        start_str = start_date.strftime('%d-%b-%Y')
        end_str = end_date.strftime('%d-%b-%Y')
        search_criteria.append(f'SINCE {start_str}')
        search_criteria.append(f'BEFORE {end_str}')

    if args.regex:
        pass  # Il filtro per la regex sarà applicato successivamente

    # Costruisce la stringa di ricerca per il comando SEARCH
    search_command = 'ALL'
    if search_criteria:
        search_command = ' '.join(search_criteria)

    res, messages = imap.search(None, search_command)
    if res != 'OK':
        print('Errore nella ricerca dei messaggi.')
        imap.logout()
        sys.exit(1)

    msg_ids = messages[0].split()

    print('Ricerca dei messaggi in corso...')
    total_msgs = len(msg_ids)
    filtered_msgs = []
    non_matching = 0

    for idx, msg_id in enumerate(msg_ids, 1):
        try:
            res, msg_data = imap.fetch(msg_id, '(RFC822.HEADER)')
            if res != 'OK' or not msg_data or msg_data[0] is None:
                print(f"\nWarning: Impossibile recuperare l'header completo per il messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs})")
                # Tentiamo di recuperare gli header disponibili
                try:
                    res, msg_data = imap.fetch(msg_id, '(BODY[HEADER.FIELDS (FROM TO SUBJECT DATE)])')
                    if res == 'OK' and msg_data and msg_data[0] is not None:
                        print("Header disponibili:")
                        header_data = msg_data[0][1].decode('utf-8', errors='ignore')
                        print(header_data)
                    else:
                        print("Impossibile recuperare gli header di base.")
                except Exception as e:
                    print(f"Errore nel tentativo di recuperare gli header di base: {str(e)}")
                continue

            header_data = msg_data[0][1]
            if header_data is None:
                print(f"\nWarning: Dati dell'header mancanti per il messaggio ID {msg_id}")
                continue

            try:
                header_data = header_data.decode('utf-8', errors='ignore')
            except AttributeError:
                print(f"\nWarning: Dati dell'header mancanti per il messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs})")
                continue

            # Verifica le condizioni AND
            and_match = all(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                            for header, regex in and_headers)

            # Verifica le condizioni OR
            or_match = any(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                        for header, regex in or_headers) if or_headers else True

            subject = get_header_value(header_data, 'Subject')

            if (and_match and or_match) and (not args.regex or re.search(args.regex, subject, re.IGNORECASE)):
                filtered_msgs.append((msg_id, subject,get_header_value(header_data, 'Date')))
            else:
                non_matching += 1

        except Exception as e:
            print(f"\nErrore durante l'elaborazione del messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs}): {str(e)}")
            continue

        # Aggiorna la barra di avanzamento
        matching = len(filtered_msgs)
        progress_bar = create_progress_bar(total_msgs, idx, matching, non_matching)
        print(f'\r{progress_bar}', end='', flush=True)
        # Fine ciclo for
        
        
    print('\n')  # Nuova linea dopo la barra di avanzamento
    num_msgs = len(filtered_msgs)
    print(f'Numero di messaggi trovati: {num_msgs}')

    if num_msgs == 0:
        print('Nessun messaggio corrisponde ai criteri di ricerca.')
        imap.logout()
        sys.exit(0)

    messages_to_delete = show_grouped_subjects_and_select(filtered_msgs)

    if messages_to_delete:
        action = "archiviazione" if args.archive or args.archive_to_disk else "cancellazione"
    confirm_action = input(f"Vuoi procedere con l'{action} di {len(messages_to_delete)} messaggi? (s/n): ")
    if confirm_action.lower() != 's':
        print('Operazione annullata.')
        imap.logout()
        sys.exit(0)

    if args.archive or args.archive_to_disk:
        archive_dest = args.archive or args.archive_to_disk
        print(f"Archiviazione dei messaggi in {archive_dest}...")
        for idx, msg_id in enumerate(messages_to_delete, 1):
            if args.archive:
                success = archive_message_imap(imap, msg_id, args.archive, args.folder, args.user, args.debug)
            else:
                success = archive_message_disk(msg_id, imap, args.archive_to_disk, args.folder, args.debug)
            
            if success:
                if args.archive:
                    imap.store(msg_id, '+FLAGS', r'(\Deleted)')
                elif args.expunge:
                    imap.store(msg_id, '+FLAGS', r'(\Deleted)')
                else:
                    res = imap.copy(msg_id, 'Trash')
                    if res[0] == 'OK':
                        imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            else:
                print (f"messaggio id {msg_id.decode()} non trasferito.")
             
            if idx % 10 == 0 or idx == len(messages_to_delete):
                print(f'{idx}/{len(messages_to_delete)} messaggi elaborati.')
        imap.expunge()
        print('Archiviazione completata.')
    else:
        print('Cancellazione in corso...')
        for idx, msg_id in enumerate(messages_to_delete, 1):
            if args.expunge:
                imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            else:
                res = imap.copy(msg_id, 'Trash')
                if res[0] == 'OK':
                    imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            if idx % 10 == 0 or idx == len(messages_to_delete):
                print(f'{idx}/{len(messages_to_delete)} messaggi elaborati.')
        imap.expunge()
        print('Cancellazione completata.')
        total_msgs = len(messages_to_delete)

        
        
    imap.logout()

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == '--help':
        print_help()
    else:
        main()
        pass
