#!/usr/bin/env python3

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
        else:
            print("Parametro non riconosciuto. Prova con -u, -s, -p, -l, -f, -d, o -e.")
        print("\nInserisci un altro parametro o 'q' per uscire:")


def decode_mime_words(s):
    return ''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in decode_header(s)
    )




def show_grouped_subjects_and_select(filtered_msgs):
    subject_count = {}
    for msg_id, subject in filtered_msgs:
        if subject not in subject_count:
            subject_count[subject] = {'count': 0, 'ids': []}
        subject_count[subject]['count'] += 1
        subject_count[subject]['ids'].append(msg_id)
    
    print(f"\nSoggetti dei messaggi trovati (totale: {len(filtered_msgs)}):")
    subjects_list = sorted(subject_count.items(), key=lambda x: x[1]['count'], reverse=True)
    for i, (subject, data) in enumerate(subjects_list, 1):
        print(f"{i}. {subject} ({data['count']} messaggi)")
    
    selected_groups = []
    while True:
        choice = input("\nInserisci i numeri dei gruppi da cancellare (separati da virgola), 'a' per tutti, o 'n' per nessuno: ").lower()
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
        return decode_mime_words(match.group(1).strip())
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
                filtered_msgs.append((msg_id, subject))
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
        confirm_delete = input(f"Vuoi procedere con la cancellazione di {len(messages_to_delete)} messaggi? (s/n): ")
        if confirm_delete.lower() != 's':
            print('Operazione annullata.')
            imap.logout()
            sys.exit(0)

        total_msgs = len(messages_to_delete)
        print('Cancellazione in corso...')
        for idx, msg_id in enumerate(messages_to_delete, 1):
            if args.expunge:
                # Cancella definitivamente
                imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            else:
                # Sposta nel Cestino
                res = imap.copy(msg_id, 'Trash')
                if res[0] == 'OK':
                    imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            if idx % 10 == 0 or idx == total_msgs:
                print(f'{idx}/{total_msgs} messaggi elaborati.')

        imap.expunge()
        print('Cancellazione completata.')
    else:
        print('Nessun messaggio cancellato.')
        
    imap.logout()

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == '--help':
        print_help()
    else:
        main()
        pass
