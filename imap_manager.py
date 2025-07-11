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
Data: 2/10/2024 ultima modifica 11/07/2025
Versione: 1.1
* 11/07/25 - Aggiunta funzionalità estesa per l'opzione "-l" di lista delle cartelle imap con informazioni statistiche salvate in csv e filtri
Licenza: eredita le licenze delle librerie utilizzate, la mia parte è frutto di
         collaborazioni multiple quindi GPL
"""

import argparse
import imaplib
import re
import getpass
import sys
from datetime import datetime, timezone
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
            print(
                "-p: Password per l'accesso. Se omessa, verrà richiesta in modo sicuro durante l'esecuzione.")
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
        elif user_input == '--no-save-csv':
            print("--no-save-csv: Se presente, impedisce il salvataggio automatico del CSV delle cartelle. "
                  "Per default, quando si usa -l, viene creato un file mailboxlist-<user@server>-<timestamp>.csv")
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
                print(
                    f"DEBUG: Verifica/creazione del percorso: {current_path}")

            # Verifica se la cartella esiste
            res, folders = imap.list(
                directory=f'"{current_path}"', pattern='*')
            if debug:
                print(f"DEBUG: Risultato list per {current_path}: res={res}")

            # Gestisci il caso in cui 'folders' è None o [None]
            if res != 'OK' or not folders or folders == [None]:
                if debug:
                    print(
                        f"DEBUG: La cartella {current_path} non esiste, tentativo di creazione...")
                # Se la cartella non esiste, la creiamo
                res, data = imap.create(current_path)
                if debug:
                    print(
                        f"DEBUG: Risultato create per {current_path}: res={res}, data={data}")
                if res != 'OK':
                    print(
                        f"Errore nella creazione della cartella {current_path}")
                    return False
            else:
                if debug:
                    print(f"DEBUG: La cartella {current_path} esiste già")
                    print(f"DEBUG: {folders} {type(folders)}")
                # Verifica se la cartella ha il flag \\HasChildren
                # Filtra eventuali valori None in 'folders'
                valid_folders = [
                    folder for folder in folders if folder is not None]
                if not any(b'\\HasChildren' in folder for folder in valid_folders):
                    if debug:
                        print(
                            f"DEBUG: La cartella {current_path} non ha il flag \\HasChildren, aggiornamento...")
                    # Aggiorna i flag della cartella (se necessario)
                    # Nota: potrebbe essere necessario un comando specifico per aggiornare i flag
    except imaplib.IMAP4.error as e:
        print(
            f"Errore IMAP durante la creazione della cartella {folder_name}: {str(e)}")
        return False
    if debug:
        print(
            f"DEBUG: Creazione della cartella {folder_name} completata con successo")
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
            print(
                f"DEBUG: Fallimento nella creazione della cartella {archive_path}")
        return False

    # Verifica se la cartella esiste prima di tentare l'append
    res, _ = imap.list(f'"{archive_path}"')
    if res != 'OK':
        print(
            f"Errore: La cartella {archive_path} non esiste o non è accessibile")
        return False

    if debug:
        print(f"DEBUG: Tentativo di append del messaggio in {archive_path}")
    #res = imap.append(f'"{archive_path}"', '', imaplib.Time2Internaldate(time.time()), email_body)
    res = imap.append(archive_path, '', imaplib.Time2Internaldate(
        time.time()), email_body)

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
        archive_path = os.path.join(
            dest_folder, source_folder_name, year, month)
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
    subjects_list = sorted(subject_count.items(),
                           key=lambda x: x[1]['count'], reverse=True)
    for i, (subject, data) in enumerate(subjects_list, 1):
        dates = sorted(data['dates'])
        if len(dates) > 1:
            date_info = f"dal {dates[0]} al {dates[-1]}"
        else:
            date_info = f"il {dates[0]}"
        print(f"{i}. {subject} ({data['count']} messaggi) - {date_info}")

    selected_groups = []
    while True:
        choice = input(
            "\nInserisci i numeri dei gruppi da selezionare (separati da virgola), 'a' per tutti, o 'n' per nessuno: ").lower()
        if choice == 'n':
            break
        elif choice == 'a':
            selected_groups = list(range(1, len(subjects_list) + 1))
            break
        else:
            try:
                selected_groups = [int(x.strip())
                                   for x in choice.split(',') if x.strip()]
                if all(1 <= x <= len(subjects_list) for x in selected_groups):
                    break
                else:
                    print("Alcuni numeri non sono validi. Riprova.")
            except ValueError:
                print(
                    "Input non valido. Inserisci numeri separati da virgola, 'a' o 'n'.")

    messages_to_delete = []
    for i in selected_groups:
        subject, data = subjects_list[i-1]
        messages_to_delete.extend(data['ids'])
        print(
            f"Selezionato per la cancellazione: {subject} ({data['count']} messaggi)")

    return messages_to_delete


def create_progress_bar(total, current, matching, non_matching):
    width, _ = shutil.get_terminal_size()

    # Calcola lo spazio necessario per la percentuale e i caratteri accessori
    percent = f" {current/total*100:.1f}%"
    extra_chars = 2  # Per le parentesi quadre []

    # Sottrai lo spazio per la percentuale e i caratteri accessori
    available_width = width - len(percent) - extra_chars

    if available_width <= 0:
        return f"[{'█' * matching}{'▒' * non_matching}{' ' * (total-current)}]{percent}"

    filled = int(available_width * current // total)
    matching_width = int(available_width * matching // total)
    non_matching_width = filled - matching_width
    remaining_width = available_width - filled

    # Usa caratteri semplici per la barra
    matching_str = '█' * matching_width
    non_matching_str = '▒' * non_matching_width
    remaining_str = ' ' * remaining_width

    bar = matching_str + non_matching_str + remaining_str
    return f"[{bar}]{percent}"


def parse_args():
    parser = argparse.ArgumentParser(
        description='Script IMAP per gestione messaggi.')
    parser.add_argument('-l', '--list', action='store_true',
                        help='Elenca le cartelle IMAP disponibili.')
    parser.add_argument('-f', '--folder', metavar='NOMECARTELLA',
                        help='Nome della cartella IMAP da selezionare.')
    parser.add_argument(
        '-s', '--server', metavar='SERVER[:PORTA]', required=True, help='Server IMAP e porta (opzionale).')
    parser.add_argument('-u', '--user', metavar='USERNAME',
                        required=True, help='Username per l\'autenticazione.')
    parser.add_argument('-p', '--password', metavar='PASSWORD',
                        nargs='?', help='Password per l\'autenticazione.')
    parser.add_argument('-d', '--datascope', metavar='INIZIO-FINE',
                        help='Intervallo di date nel formato dd/mm/yyyy-dd/mm/yyyy.')
    parser.add_argument('-e', '--expunge', action='store_true',
                        help='Cancella definitivamente i messaggi.')
    parser.add_argument(
        'regex', nargs='?', help='Espressione regolare per filtrare i soggetti dei messaggi.')
    parser.add_argument('-a', '--and-header', action='append', nargs=2, metavar=('HEADER', 'REGEX'),
                        help='Ricerca AND nell\'header specificato usando la regex fornita')
    parser.add_argument('-o', '--or-header', action='append', nargs=2, metavar=('HEADER', 'REGEX'),
                        help='Ricerca OR nell\'header specificato usando la regex fornita')
    parser.add_argument('--archive', metavar='CARTELLA_DESTINAZIONE',
                        help='Archivia spostandoli i messaggi nella cartella specificata')
    parser.add_argument('--archive-to-disk', metavar='CARTELLA_LOCALE_DESTINAZIONE',
                        help='Archivia i messaggi nella cartella locale specificata')
    parser.add_argument('--debug', action='store_true',
                        help='Abilita i messaggi di debug')
    parser.add_argument('--no-save-csv', action='store_true',
                        help='Non salva automaticamente il CSV delle cartelle in un file')
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


def get_message_counts_by_flags(imap, folder_name):
    """Conta i messaggi per ogni flag IMAP"""
    try:
        select_result, _ = imap.select(f'"{folder_name}"', readonly=True)
        if select_result != 'OK':
            return {'ERROR': 'Cannot select folder'}

        counts = {}

        # Conta tutti i messaggi
        res, messages = imap.search(None, 'ALL')
        counts['TOTAL'] = len(
            messages[0].split()) if res == 'OK' and messages[0] else 0

        # Conta messaggi per flag specifici
        flag_searches = {
            'SEEN': 'SEEN',
            'UNSEEN': 'UNSEEN',
            'ANSWERED': 'ANSWERED',
            'UNANSWERED': 'UNANSWERED',
            'FLAGGED': 'FLAGGED',
            'UNFLAGGED': 'UNFLAGGED',
            'DELETED': 'DELETED',
            'UNDELETED': 'UNDELETED',
            'DRAFT': 'DRAFT',
            'UNDRAFT': 'UNDRAFT',
            'RECENT': 'RECENT'
        }

        for flag_name, search_term in flag_searches.items():
            try:
                res, messages = imap.search(None, search_term)
                counts[flag_name] = len(
                    messages[0].split()) if res == 'OK' and messages[0] else 0
            except:
                counts[flag_name] = 0

        return counts
    except Exception as e:
        return {'ERROR': str(e)}


def list_folders(imap, args=None):
    import csv
    import sys
    import time
    from io import StringIO
    from email.parser import Parser

    result, folders = imap.list()
    if result != 'OK':
        print('Impossibile recuperare le cartelle.')
        return

    # AGGIUNTA: Filtra le cartelle se è specificato un filtro
    if args and args.folder:
        filtered_folders = []
        filter_folder = args.folder.strip('"')

        for folder in folders:
            # Parsing del nome della cartella
            parts = folder.decode().split(' "/" ')
            if len(parts) == 2:
                folder_name = parts[1].strip('"')
            else:
                folder_decoded = folder.decode()
                folder_name = folder_decoded.split(
                    '"')[-2] if '"' in folder_decoded else folder_decoded

            # Controlla se la cartella corrisponde al filtro
            if folder_name == filter_folder or folder_name.startswith(filter_folder + '/'):
                filtered_folders.append(folder)

        folders = filtered_folders

        if not folders:
            print(
                f'Nessuna cartella trovata che corrisponde al filtro: {filter_folder}')
            return

    # Prepara i criteri di ricerca se specificati
    search_criteria = []
    if args and args.datascope:
        start_date, end_date = get_date_range(args.datascope)
        start_str = start_date.strftime('%d-%b-%Y')
        end_str = end_date.strftime('%d-%b-%Y')
        search_criteria.append(f'SINCE {start_str}')
        search_criteria.append(f'BEFORE {end_str}')

    search_command = 'ALL'
    if search_criteria:
        search_command = ' '.join(search_criteria)

    and_headers = args.and_header or [] if args else []
    or_headers = args.or_header or [] if args else []

    # OTTIMIZZAZIONE: Pre-compila le regex
    compiled_and_headers = [(header, re.compile(regex, re.IGNORECASE))
                            for header, regex in and_headers] if and_headers else []
    compiled_or_headers = [(header, re.compile(regex, re.IGNORECASE))
                           for header, regex in or_headers] if or_headers else []
    compiled_subject_regex = re.compile(
        args.regex, re.IGNORECASE) if args and args.regex else None

    print("Fase 1: Scansione delle cartelle e conteggio messaggi...")

    # PRIMA FASE: Scansione rapida per contare i messaggi totali
    folder_info = []
    total_messages_to_process = 0
    start_time_phase1 = time.time()

    for folder_idx, folder in enumerate(folders, 1):
        # Parsing del nome della cartella
        parts = folder.decode().split(' "/" ')
        if len(parts) == 2:
            folder_name = parts[1].strip('"')
            folder_flags = parts[0].strip('()').split()
        else:
            folder_decoded = folder.decode()
            folder_name = folder_decoded.split(
                '"')[-2] if '"' in folder_decoded else folder_decoded
            folder_flags = []

        # Determina il tipo di cartella
        if '\\Noselect' in folder_flags:
            folder_type = "CONTAINER"
            msg_count = 0
        else:
            folder_type = "MAILBOX"
            msg_count = 0

            # Conta i messaggi nella cartella
            try:
                select_result, select_data = imap.select(
                    f'"{folder_name}"', readonly=True)
                if select_result == 'OK':
                    res, messages = imap.search(None, search_command)
                    if res == 'OK' and messages[0]:
                        msg_count = len(messages[0].split())
                    total_messages_to_process += msg_count
                else:
                    # Prova con STATUS se SELECT non funziona
                    try:
                        status_result, status_data = imap.status(
                            f'"{folder_name}"', '(MESSAGES)')
                        if status_result == 'OK' and status_data:
                            status_str = status_data[0].decode()
                            match = re.search(r'MESSAGES (\d+)', status_str)
                            if match:
                                msg_count = int(match.group(1))
                                total_messages_to_process += msg_count
                    except:
                        msg_count = "ERR"
            except Exception as e:
                if args and args.debug:
                    print(
                        f"\nDEBUG: Errore nel conteggio per {folder_name}: {str(e)}")
                msg_count = "ERR"

        folder_info.append({
            'name': folder_name,
            'flags': folder_flags,
            'type': folder_type,
            'estimated_msgs': msg_count
        })

        # Mostra progresso scansione con ETA
        elapsed_time = time.time() - start_time_phase1
        if elapsed_time > 0:
            avg_time_per_folder = elapsed_time / folder_idx
            remaining_folders = len(folders) - folder_idx
            eta_seconds = remaining_folders * avg_time_per_folder
            eta_str = f" - ETA: {eta_seconds:.1f}s" if eta_seconds > 0 else ""
        else:
            eta_str = ""

        # Tronca il nome della cartella se troppo lungo
        display_name = folder_name if len(
            folder_name) <= 40 else folder_name[:37] + "..."
        print(
            f"\rScansione: {folder_idx}/{len(folders)} - {display_name}: {msg_count} msg{eta_str}", end='', flush=True)

    print(f"\n\nTotale messaggi da elaborare: {total_messages_to_process}")
    print("\nElenco cartelle trovate:")
    for info in folder_info:
        print(
            f"  {info['estimated_msgs']:>6} msg - {info['name']} ({info['type']})")

    print(f"\nFase 2: Elaborazione dettagliata...")

    # MOSTRA I FILTRI ATTIVI
    active_filters = []
    command_line_parts = []

    # Ricostruisci la command line
    if args:
        command_line_parts.append(f"imap_manager.py")
        command_line_parts.append(f"-u {args.user}")
        command_line_parts.append(f"-s {args.server}")
        if args.folder:
            command_line_parts.append(f'-f "{args.folder}"')
        if args.list:
            command_line_parts.append("-l")
        if args.regex:
            command_line_parts.append(f'"{args.regex}"')
        if args.datascope:
            command_line_parts.append(f"-d {args.datascope}")
        if args.expunge:
            command_line_parts.append("-e")
        if args.debug:
            command_line_parts.append("--debug")
        if args.no_save_csv:
            command_line_parts.append("--no-save-csv")

    command_line = " ".join(command_line_parts)

    # Aggiungi tutti i filtri attivi
    if args and args.folder:
        active_filters.append(f"Filtro cartella: '{args.folder}'")
    if args and args.regex:
        active_filters.append(f"Regex oggetto: '{args.regex}'")
    if args and args.datascope:
        active_filters.append(f"Intervallo date: {args.datascope}")
    if and_headers:
        for header, regex in and_headers:
            active_filters.append(f"AND {header}: '{regex}'")
    if or_headers:
        for header, regex in or_headers:
            active_filters.append(f"OR {header}: '{regex}'")

    if active_filters:
        print("🔍 FILTRI ATTIVI:")
        for filter_desc in active_filters:
            print(f"   - {filter_desc}")
        print(f"   → Solo i messaggi che rispettano TUTTI i filtri verranno mostrati nei risultati\n")
    else:
        print("ℹ️  Nessun filtro attivo - verranno mostrati tutti i messaggi\n")

    # SECONDA FASE: Elaborazione dettagliata con ETA ottimizzato
    csv_data = []

    # AGGIUNGI: Righe di intestazione con informazioni sui filtri e comando
    csv_data.append(
        [f"# IMAP Manager Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
    csv_data.append([f"# Command: {command_line}"])
    csv_data.append([f"# Server: {args.server}"])
    csv_data.append([f"# User: {args.user}"])
    if active_filters:
        csv_data.append([f"# Filtri attivi: {len(active_filters)}"])
        for filter_desc in active_filters:
            csv_data.append([f"# - {filter_desc}"])
    else:
        csv_data.append([f"# Nessun filtro attivo"])

    # Stampa le informazioni sui filtri e l'header CSV
    for row in csv_data:
        if len(row) == 1 and row[0].startswith('#'):
            print(f"# {row[0][2:]}")  # Stampa i commenti senza il prefisso #
        elif len(row) == 13:  # È la riga dell'header
            print(','.join(row))
            break

    csv_data.append(['Folder', 'Total_Messages', 'Seen', 'Unseen', 'Answered', 'Flagged',
                     'Deleted', 'Draft', 'Recent', 'Oldest_Date', 'Newest_Date', 'Size_Bytes', 'Type'])

    # Variabili per ETA a campione ogni 500 messaggi
    processed_messages = 0
    processed_folders = 0
    start_time = time.time()
    current_eta = None
    last_eta_calculation = 0
    ETA_SAMPLE_INTERVAL = 500  # Calcola ETA ogni 500 messaggi
    eta_history = []  # Storico per media mobile

    # Stampa l'header CSV
    print('Folder,Total_Messages,Seen,Unseen,Answered,Flagged,Deleted,Draft,Recent,Oldest_Date,Newest_Date,Size_Bytes,Type')

    for folder_idx, info in enumerate(folder_info, 1):
        folder_name = info['name']
        folder_flags = info['flags']
        folder_type = info['type']
        estimated_msgs = info['estimated_msgs']

        # Calcola ETA per le cartelle
        if folder_idx > 1:
            current_time = time.time()
            elapsed = current_time - start_time
            avg_time_per_folder = elapsed / (folder_idx - 1)
            remaining_folders = len(folder_info) - folder_idx
            folder_eta_seconds = remaining_folders * avg_time_per_folder
            folder_eta_str = f" - ETA cartelle: {folder_eta_seconds/60:.1f}m" if folder_eta_seconds > 60 else f" - ETA: {folder_eta_seconds:.1f}s"
        else:
            folder_eta_str = ""

        # Mostra progresso cartelle
        display_name = folder_name if len(
            folder_name) <= 30 else folder_name[:27] + "..."
        filter_suffix = " (FILTRATO)" if (
            (args and args.regex) or and_headers or or_headers) else ""
        print(
            f"\rElaborando cartella {folder_idx}/{len(folder_info)}: {display_name} ({estimated_msgs} msg{filter_suffix}){folder_eta_str}", end='', flush=True)

        # Ottieni sempre i conteggi per flag, indipendentemente dal tipo di cartella
        flag_counts = get_message_counts_by_flags(imap, folder_name)

        # Determina il tipo di cartella basato sui flag
        if '\\Noselect' in folder_flags:
            # Cartella che non può contenere messaggi, solo sottocartelle
            folder_type = "CONTAINER"
        elif '\\HasChildren' in folder_flags and flag_counts.get('TOTAL', 0) == 0:
            # Cartella con sottocartelle ma senza messaggi
            folder_type = "PARENT"
        elif '\\HasChildren' in folder_flags and flag_counts.get('TOTAL', 0) > 0:
            # Cartella con sottocartelle E messaggi
            folder_type = "PARENT+MAILBOX"
        else:
            # Cartella normale con messaggi
            folder_type = "MAILBOX"

        # Inizializza valori di default prima dell'elaborazione
        msg_count = flag_counts.get('TOTAL', 0)
        oldest_date = "N/A"
        newest_date = "N/A"
        folder_size = "N/A"

        # Solo se ci sono messaggi da elaborare, procedi con l'analisi dettagliata
        if msg_count > 0:
            # INIZIALIZZA SEMPRE filtered_msg_ids
            filtered_msg_ids = []

            try:
                select_result, select_data = imap.select(
                    f'"{folder_name}"', readonly=True)
                if select_result == 'OK':
                    # Usa la ricerca con i criteri specificati (se presenti)
                    res, messages = imap.search(None, search_command)
                    if res == 'OK' and messages[0]:
                        msg_ids = messages[0].split()
                    else:
                        msg_ids = []

                    if msg_ids:
                        # Solo se ci sono criteri di filtro aggiuntivi, elabora i messaggi singolarmente
                        if (args and args.regex) or and_headers or or_headers:
                            # Elaborazione con filtri
                            total_size = 0
                            dates = []

                            # AGGIUNGI: Contatori per i flag dei messaggi filtrati
                            filtered_flag_counts = {
                                'SEEN': 0,
                                'UNSEEN': 0,
                                'ANSWERED': 0,
                                'FLAGGED': 0,
                                'DELETED': 0,
                                'DRAFT': 0,
                                'RECENT': 0
                            }

                            # Elaborazione dettagliata dei messaggi con filtri
                            for msg_idx, msg_id in enumerate(msg_ids, 1):
                                try:
                                    # Aggiorna il progresso ogni 100 messaggi
                                    if processed_messages % 100 == 0:
                                        if processed_messages >= last_eta_calculation + ETA_SAMPLE_INTERVAL:
                                            current_time = time.time()
                                            elapsed = current_time - start_time
                                            if processed_messages > 0:
                                                avg_time_per_msg = elapsed / processed_messages
                                                remaining_msgs = total_messages_to_process - processed_messages
                                                eta_seconds = remaining_msgs * avg_time_per_msg

                                                # Media mobile per ETA più stabile
                                                eta_history.append(eta_seconds)
                                                if len(eta_history) > 5:
                                                    eta_history.pop(0)
                                                current_eta = sum(
                                                    eta_history) / len(eta_history)

                                                last_eta_calculation = processed_messages

                                            eta_str = f" - ETA: {current_eta/60:.1f}m" if current_eta else ""
                                            print(
                                                f"\rElaborando {folder_name}: {msg_idx}/{len(msg_ids)} ({processed_messages}/{total_messages_to_process}){eta_str}", end='', flush=True)

                                    # MODIFICA: Recupera header, dimensione E flag del messaggio
                                    res, msg_data = imap.fetch(
                                        msg_id, '(RFC822.HEADER RFC822.SIZE FLAGS)')
                                    if res != 'OK' or not msg_data or msg_data[0] is None:
                                        processed_messages += 1
                                        continue

                                    # DEBUG: Stampa la struttura della risposta IMAP (solo per i primi 3 messaggi)
                                    if args and args.debug and msg_idx <= 3:
                                        print(
                                            f"\nDEBUG: Struttura completa msg_data per {msg_id}:")
                                        print(f"  - Tipo: {type(msg_data)}")
                                        print(
                                            f"  - Lunghezza: {len(msg_data)}")
                                        for i, item in enumerate(msg_data):
                                            print(
                                                f"  - msg_data[{i}]: {type(item)} - {str(item)[:200]}...")

                                    header_data = msg_data[0][1]
                                    if header_data is None:
                                        processed_messages += 1
                                        continue

                                    try:
                                        header_data = header_data.decode(
                                            'utf-8', errors='ignore')
                                    except AttributeError:
                                        processed_messages += 1
                                        continue

                                    # Verifica le condizioni AND
                                    and_match = all(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                                                    for header, regex in and_headers) if and_headers else True

                                    # Verifica le condizioni OR
                                    or_match = any(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                                                   for header, regex in or_headers) if or_headers else True

                                    # Verifica il subject con regex
                                    subject = get_header_value(
                                        header_data, 'Subject')
                                    regex_match = not args.regex or re.search(
                                        args.regex, subject, re.IGNORECASE)

                                    # Se il messaggio corrisponde ai criteri
                                    if and_match and or_match and regex_match:
                                        filtered_msg_ids.append(msg_id)

                                        # AGGIUNGI: Estrai e conta i flag del messaggio filtrato
                                        try:
                                            # I flag si trovano nella risposta del fetch
                                            if msg_data and msg_data[0] and len(msg_data[0]) > 0:
                                                response_line = msg_data[0][0]
                                                if isinstance(response_line, bytes):
                                                    response_line = response_line.decode(
                                                        'utf-8', errors='ignore')
                                                else:
                                                    response_line = str(
                                                        response_line)

                                                # Cerca i flag nella risposta (FLAGS (\Seen \Answered ...))
                                                flag_match = re.search(
                                                    r'FLAGS \(([^)]*)\)', response_line)
                                                if flag_match:
                                                    flags_str = flag_match.group(
                                                        1)
                                                    if args and args.debug and msg_idx <= 10:
                                                        print(
                                                            f"\nDEBUG: Messaggio {msg_id} flag: {flags_str}")

                                                    # Conta i flag
                                                    if '\\Seen' in flags_str:
                                                        filtered_flag_counts['SEEN'] += 1
                                                    else:
                                                        filtered_flag_counts['UNSEEN'] += 1

                                                    if '\\Answered' in flags_str:
                                                        filtered_flag_counts['ANSWERED'] += 1

                                                    if '\\Flagged' in flags_str:
                                                        filtered_flag_counts['FLAGGED'] += 1

                                                    if '\\Deleted' in flags_str:
                                                        filtered_flag_counts['DELETED'] += 1

                                                    if '\\Draft' in flags_str:
                                                        filtered_flag_counts['DRAFT'] += 1

                                                    if '\\Recent' in flags_str:
                                                        filtered_flag_counts['RECENT'] += 1
                                                else:
                                                    # Se non troviamo i flag, assumiamo che sia non letto
                                                    filtered_flag_counts['UNSEEN'] += 1
                                                    if args and args.debug and msg_idx <= 10:
                                                        print(
                                                            f"\nDEBUG: Messaggio {msg_id} - flag non trovati nella risposta")

                                        except Exception as e:
                                            if args and args.debug:
                                                print(
                                                    f"\nDEBUG: Errore estrazione flag per {msg_id}: {str(e)}")
                                            # Fallback: conta come non letto
                                            filtered_flag_counts['UNSEEN'] += 1

                                        # Calcola la dimensione del messaggio dalla risposta IMAP
                                        try:
                                            # La dimensione è nella prima parte della risposta come stringa
                                            if msg_data and msg_data[0] and len(msg_data[0]) > 0:
                                                response_line = msg_data[0][0]
                                                if isinstance(response_line, bytes):
                                                    response_line = response_line.decode(
                                                        'utf-8', errors='ignore')
                                                else:
                                                    response_line = str(
                                                        response_line)

                                                # Cerca RFC822.SIZE nella stringa di risposta
                                                size_match = re.search(
                                                    r'RFC822\.SIZE (\d+)', response_line)
                                                if size_match:
                                                    message_size = int(
                                                        size_match.group(1))
                                                    total_size += message_size
                                                    if args and args.debug and msg_idx <= 10:
                                                        print(
                                                            f"\nDEBUG: Messaggio {msg_id} dimensione: {message_size} bytes")
                                                else:
                                                    # Fallback: usa la lunghezza dell'header
                                                    total_size += len(header_data)
                                                    if args and args.debug and msg_idx <= 10:
                                                        print(
                                                            f"\nDEBUG: Messaggio {msg_id} - RFC822.SIZE non trovato, usando lunghezza header: {len(header_data)} bytes")
                                            else:
                                                # Fallback: usa la lunghezza dell'header
                                                total_size += len(header_data)
                                                if args and args.debug and msg_idx <= 10:
                                                    print(
                                                        f"\nDEBUG: Messaggio {msg_id} - Risposta IMAP vuota, usando lunghezza header: {len(header_data)} bytes")

                                        except Exception as e:
                                            if args and args.debug:
                                                print(
                                                    f"\nDEBUG: Errore calcolo dimensione per {msg_id}: {str(e)}")
                                            total_size += len(header_data)

                                        # Estrai la data del messaggio dal header originale
                                        try:
                                            # Cerca la data direttamente nell'header invece di usare get_header_value
                                            date_match = re.search(
                                                r'Date: (.*)', header_data, re.IGNORECASE | re.MULTILINE)
                                            if date_match:
                                                raw_date_str = date_match.group(
                                                    1).strip()
                                                # Rimuovi eventuali commenti dalle date (testo tra parentesi)
                                                raw_date_str = re.sub(
                                                    r'\s*\(.*?\)\s*', '', raw_date_str)
                                                try:
                                                    parsed_date = email.utils.parsedate_to_datetime(
                                                        raw_date_str)
                                                    if parsed_date.tzinfo is None:
                                                        parsed_date = parsed_date.replace(
                                                            tzinfo=timezone.utc)
                                                    dates.append(parsed_date)
                                                except Exception as parse_error:
                                                    if args and args.debug:
                                                        print(
                                                            f"\nDEBUG: Errore parsing data '{raw_date_str}': {str(parse_error)}")
                                        except Exception as e:
                                            if args and args.debug:
                                                print(
                                                    f"\nDEBUG: Errore estrazione data per {msg_id}: {str(e)}")

                                    processed_messages += 1

                                except Exception as e:
                                    if args and args.debug:
                                        print(
                                            f"\nDEBUG: Errore nell'elaborazione del messaggio {msg_id}: {str(e)}")
                                    processed_messages += 1
                                    continue

                            # AGGIORNA LE INFORMAZIONI SOLO SE CI SONO MESSAGGI FILTRATI
                            if filtered_msg_ids:
                                folder_size = total_size
                                # Calcola le date per i messaggi filtrati
                                if dates:
                                    dates.sort()
                                    oldest_date = dates[0].strftime('%Y-%m-%d')
                                    newest_date = dates[-1].strftime(
                                        '%Y-%m-%d')
                                else:
                                    oldest_date = "N/A"
                                    newest_date = "N/A"

                                # AGGIORNA: Usa i conteggi dei flag filtrati invece di flag_counts
                                flag_counts = filtered_flag_counts.copy()
                                flag_counts['TOTAL'] = len(filtered_msg_ids)

                                if args and args.debug:
                                    print(
                                        f"\nDEBUG: Cartella {folder_name} - Messaggi filtrati: {len(filtered_msg_ids)}")
                                    print(
                                        f"DEBUG: Dimensione totale: {folder_size} bytes")
                                    print(
                                        f"DEBUG: Date: {oldest_date} - {newest_date}")
                                    print(f"DEBUG: Flag counts: {flag_counts}")
                            else:
                                # Nessun messaggio corrisponde ai filtri
                                folder_size = 0
                                oldest_date = "N/A"
                                newest_date = "N/A"
                                # Azzera tutti i conteggi per i flag
                                flag_counts = {
                                    flag: 0 for flag in flag_counts.keys()}
                                flag_counts['TOTAL'] = 0
                        else:
                            # Se non ci sono filtri aggiuntivi, usa tutti i messaggi
                            filtered_msg_ids = msg_ids

                            # Calcola le date dei messaggi estremi
                            print(
                                f"\rElaborando {folder_name}: analisi rapida di {len(msg_ids)} messaggi...", end='', flush=True)
                            try:
                                # Ottieni il primo e ultimo messaggio per le date
                                first_msg = imap.fetch(
                                    msg_ids[0], '(RFC822.HEADER)')
                                last_msg = imap.fetch(
                                    msg_ids[-1], '(RFC822.HEADER)')

                                dates = []
                                for msg_data in [first_msg, last_msg]:
                                    if msg_data[0] == 'OK' and msg_data[1] and msg_data[1][0]:
                                        try:
                                            header_data = msg_data[1][0][1].decode(
                                                'utf-8', errors='ignore')
                                            parser = Parser()
                                            parsed_headers = parser.parsestr(
                                                header_data)
                                            date_header = parsed_headers.get(
                                                'Date')
                                            if date_header:
                                                parsed_date = email.utils.parsedate_to_datetime(
                                                    date_header)
                                                if parsed_date.tzinfo is None:
                                                    parsed_date = parsed_date.replace(
                                                        tzinfo=timezone.utc)
                                                dates.append(parsed_date)
                                        except:
                                            pass

                                if dates:
                                    dates.sort()
                                    oldest_date = dates[0].strftime('%Y-%m-%d')
                                    newest_date = dates[-1].strftime(
                                        '%Y-%m-%d')

                                # Calcola la dimensione totale (approssimativa)
                                try:
                                    # Campiona alcuni messaggi per stimare la dimensione media
                                    sample_size = min(10, len(msg_ids))
                                    sample_msgs = msg_ids[:sample_size]
                                    total_sample_size = 0

                                    for sample_idx, msg_id in enumerate(sample_msgs, 1):
                                        print(
                                            f"\rElaborando {folder_name}: campionamento {sample_idx}/{sample_size}...", end='', flush=True)
                                        size_data = imap.fetch(
                                            msg_id, '(RFC822.SIZE)')
                                        if size_data[0] == 'OK':
                                            size_match = re.search(
                                                r'RFC822.SIZE (\d+)', str(size_data[1][0]))
                                            if size_match:
                                                total_sample_size += int(
                                                    size_match.group(1))

                                    if total_sample_size > 0:
                                        avg_size = total_sample_size / sample_size
                                        folder_size = int(
                                            avg_size * len(msg_ids))
                                except:
                                    folder_size = "N/A"

                                # Aggiorna il contatore dei messaggi processati
                                processed_messages += msg_count

                            except:
                                oldest_date = "N/A"
                                newest_date = "N/A"
                                folder_size = "N/A"
                    else:
                        # Nessun messaggio trovato
                        filtered_msg_ids = []

                else:
                    # Se non si riesce a selezionare, prova con STATUS
                    filtered_msg_ids = []
                    try:
                        status_result, status_data = imap.status(
                            f'"{folder_name}"', '(MESSAGES)')
                        if status_result == 'OK' and status_data:
                            status_str = status_data[0].decode()
                            match = re.search(r'MESSAGES (\d+)', status_str)
                            if match:
                                msg_count = int(match.group(1))
                                # Aggiorna anche flag_counts se necessario
                                flag_counts['TOTAL'] = msg_count
                    except:
                        pass

            except Exception as e:
                filtered_msg_ids = []
                if args and args.debug:
                    print(
                        f"\nDEBUG: Errore nell'elaborazione della cartella {folder_name}: {str(e)}")
        else:
            # Se non ci sono messaggi, inizializza filtered_msg_ids come lista vuota
            filtered_msg_ids = []

        # Prepara i dati per il CSV
        folder_data = [
            folder_name,
            flag_counts.get('TOTAL', 0),
            flag_counts.get('SEEN', 0),
            flag_counts.get('UNSEEN', 0),
            flag_counts.get('ANSWERED', 0),
            flag_counts.get('FLAGGED', 0),
            flag_counts.get('DELETED', 0),
            flag_counts.get('DRAFT', 0),
            flag_counts.get('RECENT', 0),
            oldest_date,
            newest_date,
            folder_size,
            folder_type
        ]

        # Pulisci la linea di progresso e stampa la riga CSV
        print(f'\r{" " * 120}', end='')  # Pulisce la linea
        csv_line = ','.join(str(x) for x in folder_data)
        print(f'\r{csv_line}')

        # Aggiungi ai dati CSV
        csv_data.append(folder_data)
        processed_folders += 1

    # Stampa le statistiche finali
    total_time = time.time() - start_time
    if active_filters:
        print(f'\n📊 RISULTATI FILTRATI:')
        print(f'   - Cartelle elaborate: {len(folder_info)}')
        print(f'   - Messaggi totali scansionati: {total_messages_to_process}')
        print(
            f'   - Messaggi che rispettano i filtri: {sum(row[1] for row in csv_data if len(row) == 13 and row[0] != "Folder" and isinstance(row[1], int))}')
        print(f'   - Tempo elaborazione: {total_time:.1f} secondi')
    else:
        print(
            f'\nElaborazione completata in {total_time:.1f} secondi - {len(folder_info)} cartelle elaborate, {processed_messages} messaggi processati')

    # Salvataggio automatico del CSV (se non disabilitato)
    if not (args and args.no_save_csv):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        server_name = args.server.split(
            ':')[0] if args and args.server else 'unknown'
        username = args.user if args and args.user else 'unknown'
        filename = f"mailboxlist-{username.replace('@','_')}_{server_name}-{timestamp}.csv"

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)

                # Scrivi le righe di intestazione (commenti)
                for row in csv_data:
                    if len(row) == 1 and row[0].startswith('#'):
                        writer.writerow(row)
                    elif len(row) == 0:  # Riga vuota
                        writer.writerow([''])
                    elif len(row) == 13:  # Header o dati
                        writer.writerow(row)
                        break

                # Scrivi i dati delle cartelle (salta le righe di intestazione)
                data_rows = [row for row in csv_data if len(
                    row) == 13 and row[0] != 'Folder']
                for row in data_rows:
                    writer.writerow(row)

            print(f'Dati salvati in: {filename}')
        except Exception as e:
            print(f'Errore nel salvataggio del file CSV: {e}')


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
    match = re.search(f'{header_name}: (.*)', header_data,
                      re.IGNORECASE | re.MULTILINE)
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
        list_folders(imap, args)
        imap.logout()
        sys.exit(0)

    if not args.folder:
        print('Devi specificare una cartella con il parametro -f.')
        sys.exit(1)

    if args.debug:
        print("hierarchy_delimiter:"+get_hierarchy_delimiter(imap))
        result, data = imap.namespace()
        imap.debug = 4  # Livello di debug (0-5)
        if result == 'OK':
            print(f"Namespace: {data}")
            # Analizza il namespace per ottenere eventuali prefissi o delimitatori aggiuntivi
        else:
            print("Impossibile ottenere il namespace dal server IMAP.")

    # gestisce caratteri speciali nel nome folder
    folder_name = args.folder.replace('"', '')
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
                print(
                    f"\nWarning: Impossibile recuperare l'header completo per il messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs})")
                # Tentiamo di recuperare gli header disponibili
                try:
                    res, msg_data = imap.fetch(
                        msg_id, '(BODY[HEADER.FIELDS (FROM TO SUBJECT DATE)])')
                    if res == 'OK' and msg_data and msg_data[0] is not None:
                        print("Header disponibili:")
                        header_data = msg_data[0][1].decode(
                            'utf-8', errors='ignore')
                        print(header_data)
                    else:
                        print("Impossibile recuperare gli header di base.")
                except Exception as e:
                    print(
                        f"Errore nel tentativo di recuperare gli header di base: {str(e)}")
                continue

            header_data = msg_data[0][1]
            if header_data is None:
                print(
                    f"\nWarning: Dati dell'header mancanti per il messaggio ID {msg_id}")
                continue

            try:
                header_data = header_data.decode('utf-8', errors='ignore')
            except AttributeError:
                print(
                    f"\nWarning: Dati dell'header mancanti per il messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs})")
                continue

            # Verifica le condizioni AND
            and_match = all(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                            for header, regex in and_headers)

            # Verifica le condizioni OR
            or_match = any(re.search(regex, get_header_value(header_data, header), re.IGNORECASE)
                           for header, regex in or_headers) if or_headers else True

            subject = get_header_value(header_data, 'Subject')

            if (and_match and or_match) and (not args.regex or re.search(args.regex, subject, re.IGNORECASE)):
                filtered_msgs.append(
                    (msg_id, subject, get_header_value(header_data, 'Date')))
            else:
                non_matching += 1

        except Exception as e:
            print(
                f"\nErrore durante l'elaborazione del messaggio ID {msg_id.decode()} (Indice: {idx}/{total_msgs}): {str(e)}")
            continue

        # Aggiorna la barra di avanzamento
        matching = len(filtered_msgs)
        progress_bar = create_progress_bar(
            total_msgs, idx, matching, non_matching)
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
        confirm_action = input(
            f"Vuoi procedere con l'{action} di {len(messages_to_delete)} messaggi? (s/n): ")
        if confirm_action.lower() != 's':
            print('Operazione annullata.')
            imap.logout()
            sys.exit(0)

    if args.archive or args.archive_to_disk:
        archive_dest = args.archive or args.archive_to_disk
        print(f"Archiviazione dei messaggi in {archive_dest}...")
        for idx, msg_id in enumerate(messages_to_delete, 1):
            if args.archive:
                success = archive_message_imap(
                    imap, msg_id, args.archive, args.folder, args.user, args.debug)
            else:
                success = archive_message_disk(
                    msg_id, imap, args.archive_to_disk, args.folder, args.debug)

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
                print(f"messaggio id {msg_id.decode()} non trasferito.")

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
