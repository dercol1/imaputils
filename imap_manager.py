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

2. Elencare cartelle con filtri (tipo, numero messaggi):
   python imap_manager.py -u username@example.com -s imap.example.com -l --folder-type MAILBOX --message-count-min 10

3. Elencare cartelle con filtri sui messaggi:
   python imap_manager.py -u username@example.com -s imap.example.com -l --regex "spam" --filter-logic OR

4. Cercare e gestire messaggi in una cartella specifica:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -d "01/01/2023-31/12/2023" "oggetto da cercare"

5. Combinare filtri con logica AND/OR:
   python imap_manager.py -u username@example.com -s imap.example.com -f "INBOX" -a "From" "spam" -o "Subject" "offer" --filter-logic AND

IMPORTANTE: 
- Con -l (list) NON sono permesse operazioni di modifica (copy, move, archive, delete)
- Tutti i filtri sono applicabili sia in modalità list che in modalità gestione messaggi
- I filtri possono essere combinati con logica AND (default) o OR

Parametri per filtri:
--folder-type: Tipo cartella (CONTAINER, PARENT, PARENT+MAILBOX, MAILBOX)
--message-count-min/max/exact: Filtri per numero messaggi
--filter-logic: Logica combinazione filtri (AND/OR)
--regex: Filtro oggetto con espressione regolare
-a, --and-header: Filtri AND su header specifici (può essere usato più volte)
-o, --or-header: Filtri OR su header specifici (può essere usato più volte)
-d, --datascope: Filtro per intervallo di date (formato: "gg/mm/aaaa-gg/mm/aaaa")

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
--message-count-min: Filtra cartelle con almeno N messaggi
--message-count-max: Filtra cartelle con al massimo N messaggi  
--message-count-exact: Filtra cartelle con esattamente N messaggi
--folder-type: Filtra cartelle per tipo (CONTAINER, PARENT, PARENT+MAILBOX, MAILBOX)
--copy-to-maildir: Copia i messaggi selezionati in formato MAILDIR locale
--move-to-maildir: Sposta i messaggi selezionati in formato MAILDIR locale

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


def create_maildir_structure(maildir_path, debug=False):
    """
    Crea la struttura MAILDIR standard (cur/, new/, tmp/)
    """
    if debug:
        print(f"DEBUG: Creazione struttura MAILDIR in {maildir_path}")

    try:
        os.makedirs(maildir_path, exist_ok=True)
        os.makedirs(os.path.join(maildir_path, 'cur'), exist_ok=True)
        os.makedirs(os.path.join(maildir_path, 'new'), exist_ok=True)
        os.makedirs(os.path.join(maildir_path, 'tmp'), exist_ok=True)

        if debug:
            print(f"DEBUG: Struttura MAILDIR creata con successo")
        return True
    except Exception as e:
        print(f"Errore nella creazione della struttura MAILDIR: {e}")
        return False


def convert_imap_flags_to_maildir(imap_flags):
    """
    Converte i flag IMAP in suffisso MAILDIR
    """
    maildir_flags = []

    if '\\Seen' in imap_flags:
        maildir_flags.append('S')
    if '\\Answered' in imap_flags:
        maildir_flags.append('R')
    if '\\Flagged' in imap_flags:
        maildir_flags.append('F')
    if '\\Draft' in imap_flags:
        maildir_flags.append('D')
    if '\\Deleted' in imap_flags:
        maildir_flags.append('T')

    return ''.join(sorted(maildir_flags))


def copy_message_to_maildir(imap, msg_id, maildir_path, folder_name, debug=False):
    """
    Copia un singolo messaggio dalla struttura IMAP al formato MAILDIR
    """
    if debug:
        print(f"DEBUG: Copia messaggio {msg_id} in MAILDIR")

    try:
        # Recupera il messaggio completo con flag
        res, msg_data = imap.fetch(msg_id, '(RFC822 FLAGS)')
        if res != 'OK' or not msg_data or msg_data[0] is None:
            print(f"Errore nel recupero del messaggio {msg_id}")
            return False

        email_body = msg_data[0][1]

        # Estrai i flag dal primo elemento della risposta
        flags_response = msg_data[0][0]
        if isinstance(flags_response, bytes):
            flags_response = flags_response.decode('utf-8', errors='ignore')

        # Cerca i flag nella risposta
        flag_match = re.search(r'FLAGS \(([^)]*)\)', str(flags_response))
        imap_flags = flag_match.group(1) if flag_match else ''

        # Converti i flag IMAP in formato MAILDIR
        maildir_flags = convert_imap_flags_to_maildir(imap_flags)

        # Crea il nome del file MAILDIR
        import time
        timestamp = str(int(time.time()))
        hostname = os.uname().nodename if hasattr(os, 'uname') else 'localhost'

        # Determina la directory di destinazione (new/ per nuovi messaggi, cur/ per letti)
        target_dir = 'cur' if '\\Seen' in imap_flags else 'new'

        # Nome del file MAILDIR
        if target_dir == 'cur' and maildir_flags:
            filename = f"{timestamp}.{msg_id.decode()}.{hostname}:2,{maildir_flags}"
        else:
            filename = f"{timestamp}.{msg_id.decode()}.{hostname}"

        # Crea la sottocartella se necessario
        folder_maildir_path = os.path.join(
            maildir_path, folder_name.replace('/', '.'))
        if not create_maildir_structure(folder_maildir_path, debug):
            return False

        file_path = os.path.join(folder_maildir_path, target_dir, filename)

        if debug:
            print(f"DEBUG: Salvataggio in {file_path}")

        # Salva il messaggio
        with open(file_path, 'wb') as f:
            f.write(email_body)

        if debug:
            print(
                f"DEBUG: Messaggio salvato con successo con flag: {maildir_flags}")

        return True

    except Exception as e:
        print(f"Errore durante la copia del messaggio in MAILDIR: {e}")
        if debug:
            print(f"DEBUG: Errore dettagliato: {str(e)}")
        return False


def process_messages_to_maildir_multi_folder(imap, folder_messages_dict, base_maildir_path, base_folder, move_mode=False, expunge=False, debug=False):
    """
    SEZIONE ELABORAZIONE MULTI-CARTELLA MAILDIR
    Elabora messaggi da multiple cartelle mantenendo la struttura gerarchica
    """
    if debug:
        total_messages = sum(len(messages)
                             for messages in folder_messages_dict.values())
        print(
            f"DEBUG: Elaborazione di {total_messages} messaggi da {len(folder_messages_dict)} cartelle")
        print(f"DEBUG: Modalità: {'SPOSTA' if move_mode else 'COPIA'}")

    successful_copies = 0
    failed_copies = 0
    folders_created = 0

    print(f"{'Spostamento' if move_mode else 'Copia'} dei messaggi in formato MAILDIR...")

    for folder_name, messages in folder_messages_dict.items():
        print(
            f"\nElaborando cartella: {folder_name} ({len(messages)} messaggi)")

        # Calcola il path relativo rispetto alla cartella base
        if folder_name == base_folder:
            relative_path = ""
        else:
            relative_path = folder_name[len(base_folder):].lstrip('/')

        # Crea il percorso MAILDIR mantenendo la struttura
        if relative_path:
            folder_maildir_path = os.path.join(
                base_maildir_path, relative_path.replace('/', '.'))
        else:
            folder_maildir_path = base_maildir_path

        if debug:
            print(
                f"DEBUG: Percorso MAILDIR per {folder_name}: {folder_maildir_path}")

        # Crea la struttura MAILDIR per questa cartella
        if not create_maildir_structure(folder_maildir_path, debug):
            print(
                f"Errore nella creazione della struttura MAILDIR per {folder_name}")
            failed_copies += len(messages)
            continue

        folders_created += 1
        print(f"  ✓ Struttura MAILDIR creata: {folder_maildir_path}")

        # Se non ci sono messaggi, continua con la prossima cartella
        if not messages:
            print(f"  → Cartella vuota - nessun messaggio da elaborare")
            continue

        # Seleziona la cartella corrente solo se ci sono messaggi
        try:
            res, data = imap.select(f'"{folder_name}"', readonly=not move_mode)
            if res != 'OK':
                print(f"Impossibile selezionare la cartella {folder_name}")
                failed_copies += len(messages)
                continue
        except Exception as e:
            print(f"Errore nella selezione della cartella {folder_name}: {e}")
            failed_copies += len(messages)
            continue

        # Processa i messaggi in questa cartella
        for idx, msg_id in enumerate(messages, 1):
            success = copy_message_to_maildir_extended(
                imap, msg_id, folder_maildir_path, folder_name, debug)

            if success:
                successful_copies += 1

                # Se è modalità spostamento, marca il messaggio per la cancellazione
                if move_mode:
                    if expunge:
                        imap.store(msg_id, '+FLAGS', r'(\Deleted)')
                    else:
                        # Prova a copiare nel cestino prima di eliminare
                        try:
                            res = imap.copy(msg_id, 'Trash')
                            if res[0] == 'OK':
                                imap.store(msg_id, '+FLAGS', r'(\Deleted)')
                        except:
                            # Se non riesce a copiare nel cestino, elimina direttamente
                            imap.store(msg_id, '+FLAGS', r'(\Deleted)')
            else:
                failed_copies += 1

            # Mostra progresso ogni 10 messaggi
            if idx % 10 == 0 or idx == len(messages):
                print(
                    f'  {idx}/{len(messages)} messaggi elaborati per {folder_name}')

        # Esegui expunge se è modalità spostamento
        if move_mode and messages:
            imap.expunge()

    print(f"\nOperazione completata:")
    print(f"  - Cartelle MAILDIR create: {folders_created}")
    print(f"  - Messaggi copiati con successo: {successful_copies}")
    if failed_copies > 0:
        print(f"  - Messaggi falliti: {failed_copies}")

    return successful_copies, failed_copies


def copy_message_to_maildir_extended(imap, msg_id, maildir_path, folder_name, debug=False):
    """
    SEZIONE COPIA MESSAGGIO MAILDIR ESTESA
    Versione estesa della funzione di copia che gestisce meglio i percorsi
    """
    if debug:
        print(
            f"DEBUG: Copia messaggio {msg_id} da {folder_name} in {maildir_path}")

    try:
        # Recupera il messaggio completo con flag
        res, msg_data = imap.fetch(msg_id, '(RFC822 FLAGS)')
        if res != 'OK' or not msg_data or msg_data[0] is None:
            print(f"Errore nel recupero del messaggio {msg_id}")
            return False

        email_body = msg_data[0][1]

        # Estrai i flag dal primo elemento della risposta
        flags_response = msg_data[0][0]
        if isinstance(flags_response, bytes):
            flags_response = flags_response.decode('utf-8', errors='ignore')

        # Cerca i flag nella risposta
        flag_match = re.search(r'FLAGS \(([^)]*)\)', str(flags_response))
        imap_flags = flag_match.group(1) if flag_match else ''

        # Converti i flag IMAP in formato MAILDIR
        maildir_flags = convert_imap_flags_to_maildir(imap_flags)

        # Crea il nome del file MAILDIR
        import time
        timestamp = str(int(time.time()))
        hostname = os.uname().nodename if hasattr(os, 'uname') else 'localhost'

        # Determina la directory di destinazione (new/ per nuovi messaggi, cur/ per letti)
        target_dir = 'cur' if '\\Seen' in imap_flags else 'new'

        # Nome del file MAILDIR
        if target_dir == 'cur' and maildir_flags:
            filename = f"{timestamp}.{msg_id.decode()}.{hostname}:2,{maildir_flags}"
        else:
            filename = f"{timestamp}.{msg_id.decode()}.{hostname}"

        file_path = os.path.join(maildir_path, target_dir, filename)

        if debug:
            print(f"DEBUG: Salvataggio in {file_path}")

        # Salva il messaggio
        with open(file_path, 'wb') as f:
            f.write(email_body)

        if debug:
            print(
                f"DEBUG: Messaggio salvato con successo con flag: {maildir_flags}")

        return True

    except Exception as e:
        print(f"Errore durante la copia del messaggio in MAILDIR: {e}")
        if debug:
            print(f"DEBUG: Errore dettagliato: {str(e)}")
        return False


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
    parser.add_argument('--message-count-min', type=int, metavar='N',
                        help='Filtra cartelle con almeno N messaggi')
    parser.add_argument('--message-count-max', type=int, metavar='N',
                        help='Filtra cartelle con al massimo N messaggi')
    parser.add_argument('--message-count-exact', type=int, metavar='N',
                        help='Filtra cartelle con esattamente N messaggi')
    parser.add_argument('--folder-type', choices=['CONTAINER', 'PARENT', 'PARENT+MAILBOX', 'MAILBOX'],
                        help='Filtra cartelle per tipo specifico')

    parser.add_argument('--copy-to-maildir', metavar='PERCORSO_MAILDIR',
                        help='Copia i messaggi selezionati in formato MAILDIR locale')
    parser.add_argument('--move-to-maildir', metavar='PERCORSO_MAILDIR',
                        help='Sposta i messaggi selezionati in formato MAILDIR locale')

    # PARAMETRI PER LOGICA FILTRI
    parser.add_argument('--filter-logic', choices=['AND', 'OR'], default='AND',
                        help='Logica di combinazione dei filtri (default: AND)')

    return parser.parse_args()


def validate_args(args):
    """
    SEZIONE VALIDAZIONE ARGOMENTI
    Valida la compatibilità degli argomenti e rileva configurazioni non valide
    """
    errors = []

    # REGOLA 1: Con -l non sono permesse operazioni di modifica
    if args.list:
        modification_options = [
            ('--copy-to-maildir', args.copy_to_maildir),
            ('--move-to-maildir', args.move_to_maildir),
            ('--archive', args.archive),
            ('--archive-to-disk', args.archive_to_disk),
            ('--expunge', args.expunge),
            ('regex search', args.regex)
        ]

        active_modifications = [opt for opt,
                                value in modification_options if value]
        if active_modifications:
            errors.append(
                f"Con -l (list) non sono permesse operazioni di modifica: {', '.join(active_modifications)}")

    # REGOLA 2: Le opzioni di modifica richiedono una cartella specifica
    if not args.list and not args.folder:
        modification_options = [
            args.copy_to_maildir, args.move_to_maildir, args.archive, args.archive_to_disk]
        if any(modification_options) or args.regex:
            errors.append(
                "Le operazioni di ricerca e modifica richiedono una cartella specifica (-f)")

    # REGOLA 3: Non si possono usare più opzioni di destinazione contemporaneamente
    destination_options = [
        ('--copy-to-maildir', args.copy_to_maildir),
        ('--move-to-maildir', args.move_to_maildir),
        ('--archive', args.archive),
        ('--archive-to-disk', args.archive_to_disk)
    ]

    active_destinations = [opt for opt, value in destination_options if value]
    if len(active_destinations) > 1:
        errors.append(
            f"Non è possibile usare più opzioni di destinazione contemporaneamente: {', '.join(active_destinations)}")

    if errors:
        print("ERRORI DI VALIDAZIONE:")
        for error in errors:
            print(f"  ❌ {error}")
        sys.exit(1)

    return True

#########################################
# Class FilterManager


class FilterManager:
    """
    SEZIONE GESTIONE FILTRI
    Gestisce tutti i tipi di filtri applicabili alle cartelle e ai messaggi
    """

    def __init__(self, args, debug=False):
        self.args = args
        self.debug = debug
        self.folder_filters = []
        self.message_filters = []
        self.logic_mode = 'AND'  # Default: tutti i filtri devono essere soddisfatti

        # Compila i filtri in base agli argomenti
        self._compile_filters()

    def _compile_filters(self):
        """Compila tutti i filtri in base agli argomenti forniti"""
        if self.debug:
            print("DEBUG: Compilazione filtri in corso...")

        # FILTRI PER CARTELLE
        if self.args.folder_type:
            self.folder_filters.append(('folder_type', self.args.folder_type))

        # CORREGGI I CONTROLLI PER I FILTRI NUMERICI
        if self.args.message_count_min is not None:
            self.folder_filters.append(
                ('message_count_min', self.args.message_count_min))

        if self.args.message_count_max is not None:
            self.folder_filters.append(
                ('message_count_max', self.args.message_count_max))

        if self.args.message_count_exact is not None:
            self.folder_filters.append(
                ('message_count_exact', self.args.message_count_exact))

        # RESTO DEL CODICE rimane uguale...
        # FILTRI PER MESSAGGI
        if self.args.regex:
            self.message_filters.append(
                ('subject_regex', re.compile(self.args.regex, re.IGNORECASE)))

        if self.args.and_header:
            for header, regex in self.args.and_header:
                self.message_filters.append(
                    ('and_header', (header, re.compile(regex, re.IGNORECASE))))

        if self.args.or_header:
            for header, regex in self.args.or_header:
                self.message_filters.append(
                    ('or_header', (header, re.compile(regex, re.IGNORECASE))))

        # FILTRI PER DATE
        if self.args.datascope:
            start_date, end_date = get_date_range(self.args.datascope)
            self.message_filters.append(('date_range', (start_date, end_date)))

        if self.debug:
            print(
                f"DEBUG: Filtri cartelle compilati: {len(self.folder_filters)}")
            print(
                f"DEBUG: Filtri messaggi compilati: {len(self.message_filters)}")

    def apply_folder_filters(self, folder_info, flag_counts):
        """
        Applica i filtri a una cartella specifica
        Returns: True se la cartella passa tutti i filtri, False altrimenti
        """
        folder_name = folder_info['name']
        folder_type = folder_info['type']

        for filter_type, filter_value in self.folder_filters:
            if filter_type == 'folder_type':
                if folder_type != filter_value:
                    if self.debug:
                        print(
                            f"DEBUG: Cartella {folder_name} esclusa per tipo: {folder_type} != {filter_value}")
                    return False

            elif filter_type == 'message_count_min':
                msg_count = flag_counts.get('TOTAL', 0)
                if isinstance(msg_count, int) and msg_count < filter_value:
                    if self.debug:
                        print(
                            f"DEBUG: Cartella {folder_name} esclusa per min count: {msg_count} < {filter_value}")
                    return False

            elif filter_type == 'message_count_max':
                msg_count = flag_counts.get('TOTAL', 0)
                if isinstance(msg_count, int) and msg_count > filter_value:
                    if self.debug:
                        print(
                            f"DEBUG: Cartella {folder_name} esclusa per max count: {msg_count} > {filter_value}")
                    return False

            elif filter_type == 'message_count_exact':
                msg_count = flag_counts.get('TOTAL', 0)
                if isinstance(msg_count, int) and msg_count != filter_value:
                    if self.debug:
                        print(
                            f"DEBUG: Cartella {folder_name} esclusa per exact count: {msg_count} != {filter_value}")
                    return False

        return True

    def apply_message_filters(self, header_data, message_date=None):
        """
        Applica i filtri a un messaggio specifico
        Returns: True se il messaggio passa tutti i filtri, False altrimenti
        """
        and_results = []
        or_results = []
        general_results = []

        for filter_type, filter_value in self.message_filters:
            if filter_type == 'subject_regex':
                subject = get_header_value(header_data, 'Subject')
                match = filter_value.search(subject)
                general_results.append(bool(match))

            elif filter_type == 'and_header':
                header_name, regex = filter_value
                header_value = get_header_value(header_data, header_name)
                match = regex.search(header_value)
                and_results.append(bool(match))

            elif filter_type == 'or_header':
                header_name, regex = filter_value
                header_value = get_header_value(header_data, header_name)
                match = regex.search(header_value)
                or_results.append(bool(match))

            elif filter_type == 'date_range':
                # SKIP IL FILTRO DATE SE È GIÀ STATO APPLICATO NELLA RICERCA IMAP
                if self.debug:
                    print(
                        f"DEBUG: Saltando filtro date - già applicato in ricerca IMAP")

                continue

                # CODICE NON ESEGUITO (FALLBACK Just In CASE) - da rimuovere visto che la ricerca è già stata fatta in IMAP
                start_date, end_date = filter_value
                if message_date:
                    try:
                        if isinstance(message_date, str):
                            parsed_date = email.utils.parsedate_to_datetime(
                                message_date)
                        else:
                            parsed_date = message_date

                        if parsed_date.tzinfo is None:
                            parsed_date = parsed_date.replace(
                                tzinfo=timezone.utc)

                        # Converti le date di confronto in timezone-aware se necessario
                        if start_date.tzinfo is None:
                            start_date = start_date.replace(
                                tzinfo=timezone.utc)
                        if end_date.tzinfo is None:
                            end_date = end_date.replace(tzinfo=timezone.utc)

                        date_match = start_date <= parsed_date <= end_date
                        general_results.append(date_match)
                    except Exception as e:
                        if self.debug:
                            print(f"DEBUG: Errore parsing data: {e}")
                        general_results.append(False)
                else:
                    general_results.append(False)

        # LOGICA DI COMBINAZIONE DEI FILTRI
        # AND: tutti i filtri AND devono essere True
        and_passed = all(and_results) if and_results else True

        # OR: almeno uno dei filtri OR deve essere True
        or_passed = any(or_results) if or_results else True

        # GENERAL: tutti i filtri generali devono essere True
        general_passed = all(general_results) if general_results else True

        # Risultato finale: AND + OR + GENERAL
        final_result = and_passed and or_passed and general_passed

        if self.debug:
            print(
                f"DEBUG: Filtri messaggio - AND: {and_passed}, OR: {or_passed}, GENERAL: {general_passed}, FINALE: {final_result}")

        return final_result

    def has_message_filters(self):
        """Verifica se ci sono filtri per messaggi attivi"""
        return len(self.message_filters) > 0

    def has_folder_filters(self):
        """Verifica se ci sono filtri per cartelle attivi"""
        return len(self.folder_filters) > 0

    def get_active_filters_description(self):
        """Restituisce una descrizione dei filtri attivi"""
        descriptions = []

        for filter_type, filter_value in self.folder_filters:
            if filter_type == 'folder_type':
                descriptions.append(f"Tipo cartella: {filter_value}")
            elif filter_type == 'message_count_min':
                descriptions.append(f"Min messaggi: {filter_value}")
            elif filter_type == 'message_count_max':
                descriptions.append(f"Max messaggi: {filter_value}")
            elif filter_type == 'message_count_exact':
                descriptions.append(f"Esatto messaggi: {filter_value}")

        for filter_type, filter_value in self.message_filters:
            if filter_type == 'subject_regex':
                descriptions.append(f"Regex oggetto: '{filter_value.pattern}'")
            elif filter_type == 'and_header':
                header_name, regex = filter_value
                descriptions.append(f"AND {header_name}: '{regex.pattern}'")
            elif filter_type == 'or_header':
                header_name, regex = filter_value
                descriptions.append(f"OR {header_name}: '{regex.pattern}'")
            elif filter_type == 'date_range':
                start_date, end_date = filter_value
                descriptions.append(
                    f"Date: {start_date.strftime('%d/%m/%Y')}-{end_date.strftime('%d/%m/%Y')}")

        return descriptions

# /CLASS FilterManager
#############################################################


def find_subfolders(imap, base_folder, debug=False):
    """
    SEZIONE RICERCA SOTTOCARTELLE
    Trova tutte le sottocartelle di una cartella base
    """
    if debug:
        print(f"DEBUG: Ricerca sottocartelle per {base_folder}")

    try:
        # Ottieni tutte le cartelle
        result, folders = imap.list()
        if result != 'OK':
            print('Impossibile recuperare le cartelle.')
            return []

        subfolders = []
        base_folder = base_folder.strip('"')

        for folder in folders:
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

            # Controlla se la cartella è una sottocartella
            if folder_name == base_folder or folder_name.startswith(base_folder + '/'):
                subfolders.append({
                    'name': folder_name,
                    'flags': folder_flags,
                    'full_folder_data': folder
                })
                if debug:
                    print(f"DEBUG: Trovata sottocartella: {folder_name}")

        if debug:
            print(f"DEBUG: Trovate {len(subfolders)} sottocartelle")

        return subfolders

    except Exception as e:
        print(f"Errore nella ricerca delle sottocartelle: {e}")
        return []


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
    """
    SEZIONE ELENCO CARTELLE
    Elenca le cartelle IMAP applicando tutti i filtri specificati
    """
    import csv
    import sys
    import time
    from io import StringIO
    from email.parser import Parser

    # INIZIALIZZAZIONE FILTER MANAGER
    filter_manager = FilterManager(
        args, args.debug if args else False) if args else None

    result, folders = imap.list()
    if result != 'OK':
        print('Impossibile recuperare le cartelle.')
        return

    # FILTRO PER CARTELLA SPECIFICA (se specificata)
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

    # PREPARAZIONE CRITERI DI RICERCA
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
            folder_name) <= 40 else "..." + folder_name[-37:]
        print(
            f"\rScansione: {folder_idx}/{len(folders)} - {display_name}: {msg_count} msg{eta_str}", end='', flush=True)

    print(f"\n\nTotale messaggi da elaborare: {total_messages_to_process}")
    # MOSTRA I FILTRI ATTIVI
    if filter_manager:
        active_filters = filter_manager.get_active_filters_description()
        if active_filters:
            print("🔍 FILTRI ATTIVI:")
            for filter_desc in active_filters:
                print(f"   - {filter_desc}")
            print(f"   → Logica filtri: {args.filter_logic}")
            print(
                f"   → Solo le cartelle/messaggi che rispettano i filtri verranno mostrati\n")
        else:
            print("ℹ️  Nessun filtro attivo - verranno mostrate tutte le cartelle\n")

    print(f"\nFase 2: Elaborazione dettagliata...")

    # PREPARAZIONE DATI CSV
    csv_data = []
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Ricostruisci la command line
    command_line_parts = []

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
        if args.debug:
            command_line_parts.append("--debug")
        if args.filter_logic != 'AND':
            command_line_parts.append(f"--filter-logic {args.filter_logic}")
        if args.no_save_csv:
            command_line_parts.append("--no-save-csv")

    command_line = " ".join(command_line_parts)

    # Intestazioni CSV
    csv_data.append(
        [f"# IMAP Manager Report - {timestamp}"])
    csv_data.append([f"# Command: {command_line}"])
    csv_data.append([f"# Server: {args.server}"])
    csv_data.append([f"# User: {args.user}"])

    if filter_manager:
        active_filters = filter_manager.get_active_filters_description()
        if active_filters:
            csv_data.append([f"# Filtri attivi: {len(active_filters)}"])
            for filter_desc in active_filters:
                csv_data.append([f"# - {filter_desc}"])
        else:
            csv_data.append([f"# Nessun filtro attivo"])

    csv_data.append(['Folder', 'Total_Messages', 'Seen', 'Unseen', 'Answered', 'Flagged',
                     'Deleted', 'Draft', 'Recent', 'Oldest_Date', 'Newest_Date', 'Size_Bytes', 'Type'])

    # Stampa l'header CSV
    print('Folder,Total_Messages,Seen,Unseen,Answered,Flagged,Deleted,Draft,Recent,Oldest_Date,Newest_Date,Size_Bytes,Type')

    # ELABORAZIONE DETTAGLIATA
    processed_messages = 0
    start_time = time.time()
    folders_processed = 0
    folders_included = 0

    for folder_idx, info in enumerate(folder_info, 1):
        folder_name = info['name']
        folder_flags = info['flags']
        folder_type = info['type']
        estimated_msgs = info['estimated_msgs']

        # Ottieni sempre i conteggi per flag
        flag_counts = get_message_counts_by_flags(imap, folder_name)

        # Determina il tipo di cartella basato sui flag
        if '\\Noselect' in folder_flags:
            folder_type = "CONTAINER"
        elif '\\HasChildren' in folder_flags and flag_counts.get('TOTAL', 0) == 0:
            folder_type = "PARENT"
        elif '\\HasChildren' in folder_flags and flag_counts.get('TOTAL', 0) > 0:
            folder_type = "PARENT+MAILBOX"
        else:
            folder_type = "MAILBOX"

        # Aggiorna info del folder
        info['type'] = folder_type

        # APPLICA FILTRI PER CARTELLE
        if filter_manager and filter_manager.has_folder_filters():
            if not filter_manager.apply_folder_filters(info, flag_counts):
                continue  # Salta questa cartella

        folders_processed += 1
        msg_count = flag_counts.get('TOTAL', 0)
        oldest_date = "N/A"
        newest_date = "N/A"
        folder_size = "N/A"

        # Se ci sono filtri per messaggi, elabora i messaggi singolarmente
        if filter_manager and filter_manager.has_message_filters() and msg_count > 0:
            filtered_flag_counts = {
                'SEEN': 0, 'UNSEEN': 0, 'ANSWERED': 0, 'FLAGGED': 0,
                'DELETED': 0, 'DRAFT': 0, 'RECENT': 0, 'TOTAL': 0
            }

            try:
                select_result, select_data = imap.select(
                    f'"{folder_name}"', readonly=True)
                if select_result == 'OK':
                    res, messages = imap.search(None, search_command)
                    if res == 'OK' and messages[0]:
                        msg_ids = messages[0].split()

                        total_size = 0
                        dates = []

                        for msg_idx, msg_id in enumerate(msg_ids, 1):
                            try:
                                # Mostra progresso
                                if msg_idx % 100 == 0:
                                    print(
                                        f"\rElaborando {folder_name}: {msg_idx}/{len(msg_ids)} messaggi", end='', flush=True)

                                res, msg_data = imap.fetch(
                                    msg_id, '(RFC822.HEADER RFC822.SIZE FLAGS)')
                                if res != 'OK' or not msg_data or msg_data[0] is None:
                                    continue

                                header_data = msg_data[0][1]
                                if header_data is None:
                                    continue

                                try:
                                    header_data = header_data.decode(
                                        'utf-8', errors='ignore')
                                except AttributeError:
                                    continue

                                # Applica filtri messaggio
                                if filter_manager.apply_message_filters(header_data):
                                    filtered_flag_counts['TOTAL'] += 1

                                    # Estrai e conta i flag
                                    if msg_data and msg_data[0] and len(msg_data[0]) > 0:
                                        response_line = msg_data[0][0]
                                        if isinstance(response_line, bytes):
                                            response_line = response_line.decode(
                                                'utf-8', errors='ignore')
                                        else:
                                            response_line = str(response_line)

                                        flag_match = re.search(
                                            r'FLAGS \(([^)]*)\)', response_line)
                                        if flag_match:
                                            flags_str = flag_match.group(1)

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
                                            filtered_flag_counts['UNSEEN'] += 1

                                        # Calcola dimensione
                                        size_match = re.search(
                                            r'RFC822\.SIZE (\d+)', response_line)
                                        if size_match:
                                            message_size = int(
                                                size_match.group(1))
                                            total_size += message_size

                                    # Estrai data
                                    try:
                                        date_match = re.search(
                                            r'Date: (.*)', header_data, re.IGNORECASE | re.MULTILINE)
                                        if date_match:
                                            raw_date_str = date_match.group(
                                                1).strip()
                                            raw_date_str = re.sub(
                                                r'\s*\(.*?\)\s*', '', raw_date_str)
                                            parsed_date = email.utils.parsedate_to_datetime(
                                                raw_date_str)
                                            if parsed_date.tzinfo is None:
                                                parsed_date = parsed_date.replace(
                                                    tzinfo=timezone.utc)
                                            dates.append(parsed_date)
                                    except:
                                        pass

                                processed_messages += 1

                            except Exception as e:
                                if args and args.debug:
                                    print(
                                        f"\nDEBUG: Errore elaborazione messaggio {msg_id}: {str(e)}")
                                continue

                        # Aggiorna i conteggi con i risultati filtrati
                        flag_counts = filtered_flag_counts
                        folder_size = total_size

                        if dates:
                            dates.sort()
                            oldest_date = dates[0].strftime('%Y-%m-%d')
                            newest_date = dates[-1].strftime('%Y-%m-%d')

            except Exception as e:
                if args and args.debug:
                    print(
                        f"\nDEBUG: Errore elaborazione cartella {folder_name}: {str(e)}")

        # Solo includere cartelle che hanno messaggi (se ci sono filtri per messaggi)
        if filter_manager and filter_manager.has_message_filters():
            if flag_counts.get('TOTAL', 0) == 0:
                continue  # Salta cartelle senza messaggi che passano i filtri

        folders_included += 1

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
        print(f'\r{" " * 120}', end='')
        csv_line = ','.join(str(x) for x in folder_data)
        print(f'\r{csv_line}')

        csv_data.append(folder_data)

    # Stampa le statistiche finali
    total_time = time.time() - start_time
    print(f'\n📊 RISULTATI:')
    print(f'   - Cartelle elaborate: {folders_processed}')
    print(f'   - Cartelle incluse nei risultati: {folders_included}')
    if filter_manager and filter_manager.has_message_filters():
        print(f'   - Messaggi totali processati: {processed_messages}')
        print(
            f'   - Messaggi che rispettano i filtri: {sum(row[1] for row in csv_data if len(row) == 13 and isinstance(row[1], int))}')
    print(f'   - Tempo elaborazione: {total_time:.1f} secondi')

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
                for row in csv_data:
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
    """
    SEZIONE PRINCIPALE
    Coordina l'esecuzione di tutte le funzionalità del programma
    """

    # Parse arguments
    args = parse_args()

    # VALIDAZIONE ARGOMENTI
    validate_args(args)

    if not args.password:
        args.password = getpass.getpass('Inserisci la password: ')

    imap = connect_imap(args.server, args.user, args.password)

    # MODALITÀ ELENCO CARTELLE
    if args.list:
        list_folders(imap, args)
        imap.logout()
        sys.exit(0)

    # MODALITÀ GESTIONE MESSAGGI
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

    # SEZIONE RICERCA RICORSIVA SOTTOCARTELLE
    base_folder = args.folder.replace('"', '')
    subfolders = find_subfolders(imap, base_folder, args.debug)

    if not subfolders:
        print(f'Nessuna cartella trovata per il percorso: {base_folder}')
        imap.logout()
        sys.exit(1)

    print(f"Trovate {len(subfolders)} cartelle da elaborare:")
    for subfolder in subfolders:
        print(f"  - {subfolder['name']}")

    # INIZIALIZZAZIONE FILTER MANAGER
    filter_manager = FilterManager(args, args.debug)

    # PREPARAZIONE CRITERI DI RICERCA
    search_criteria = []
    if args.datascope:
        start_date, end_date = get_date_range(args.datascope)
        start_str = start_date.strftime('%d-%b-%Y')
        end_str = end_date.strftime('%d-%b-%Y')
        search_criteria.append(f'SINCE {start_str}')
        search_criteria.append(f'BEFORE {end_str}')

    search_command = 'ALL'
    if search_criteria:
        search_command = ' '.join(search_criteria)

    # RICERCA MESSAGGI IN TUTTE LE SOTTOCARTELLE
    folder_messages_dict = {}
    total_messages_found = 0
    folders_matching_filters = []  # NUOVO: traccia cartelle che soddisfano filtri

    print("Ricerca dei messaggi in corso...")

    for subfolder in subfolders:
        folder_name = subfolder['name']
        folder_flags = subfolder['flags']

        # Determina il tipo di cartella
        if '\\Noselect' in folder_flags:
            folder_type = "CONTAINER"
        elif '\\HasChildren' in folder_flags:
            folder_type = "PARENT"
        else:
            folder_type = "MAILBOX"

        # Applica filtri per cartelle se attivi
        folder_matches_filters = True
        if filter_manager.has_folder_filters():
            flag_counts = get_message_counts_by_flags(imap, folder_name)

            # Raffina il tipo di cartella
            if folder_type == "PARENT":
                if flag_counts.get('TOTAL', 0) > 0:
                    folder_type = "PARENT+MAILBOX"

            folder_info = {
                'name': folder_name,
                'type': folder_type,
                'flags': folder_flags
            }

            if not filter_manager.apply_folder_filters(folder_info, flag_counts):
                folder_matches_filters = False
                if args.debug:
                    print(f"DEBUG: Cartella {folder_name} esclusa dai filtri")
                continue

        # Salta cartelle container che non possono contenere messaggi
        if '\\Noselect' in folder_flags:
            if args.debug:
                print(f"DEBUG: Saltando cartella container: {folder_name}")
            continue

        # NUOVO: Aggiungi cartella che soddisfa i filtri
        if folder_matches_filters:
            folders_matching_filters.append(folder_name)

        try:
            # Seleziona la cartella
            res, data = imap.select(f'"{folder_name}"', readonly=True)
            if res != 'OK':
                print(f"Impossibile selezionare la cartella {folder_name}")
                continue

            # Cerca messaggi
            res, messages = imap.search(None, search_command)
            if res != 'OK':
                print(f"Errore nella ricerca dei messaggi in {folder_name}")
                continue

            msg_ids = messages[0].split()
            if not msg_ids:
                if args.debug:
                    print(f"DEBUG: Nessun messaggio trovato in {folder_name}")
                # NUOVO: Anche se non ci sono messaggi, aggiungi la cartella vuota
                if folder_matches_filters:
                    folder_messages_dict[folder_name] = []
                continue

            print(f"Trovati {len(msg_ids)} messaggi in {folder_name}")

            # Applica filtri sui messaggi se attivi
            if filter_manager.has_message_filters():
                filtered_msg_ids = []

                # CONTROLLA SE CI SONO FILTRI PER DATE - SE SI, SALTA IL FILTRO DATE
                # PERCHÉ GIÀ APPLICATO NELLA RICERCA IMAP
                has_date_filter = any(
                    filter_type == 'date_range' for filter_type, _ in filter_manager.message_filters)

                if args.debug:
                    print(
                        f"DEBUG: Applicazione filtri messaggio per {folder_name}")
                    print(
                        f"DEBUG: Filtri date già applicati in IMAP: {has_date_filter}")
                    print(f"DEBUG: Messaggi da filtrare: {len(msg_ids)}")

                for msg_id in msg_ids:
                    try:
                        res, msg_data = imap.fetch(msg_id, '(RFC822.HEADER)')
                        if res != 'OK' or not msg_data or msg_data[0] is None:
                            continue

                        header_data = msg_data[0][1]
                        if header_data is None:
                            continue

                        try:
                            header_data = header_data.decode(
                                'utf-8', errors='ignore')
                        except AttributeError:
                            continue

                        # ESTRAI LA DATA DAL HEADER PER PASSARLA AI FILTRI
                        message_date = None
                        if not has_date_filter:  # Solo se non c'è già un filtro date applicato
                            date_match = re.search(
                                r'Date: (.*)', header_data, re.IGNORECASE | re.MULTILINE)
                            if date_match:
                                try:
                                    raw_date_str = date_match.group(1).strip()
                                    raw_date_str = re.sub(
                                        r'\s*\(.*?\)\s*', '', raw_date_str)
                                    message_date = email.utils.parsedate_to_datetime(
                                        raw_date_str)
                                except:
                                    message_date = None

                        # APPLICA FILTRI MESSAGGIO CON LA DATA
                        if filter_manager.apply_message_filters(header_data, message_date):
                            filtered_msg_ids.append(msg_id)
                            if args.debug:
                                print(
                                    f"DEBUG: Messaggio {msg_id.decode()} passa i filtri")
                        else:
                            if args.debug:
                                print(
                                    f"DEBUG: Messaggio {msg_id.decode()} non passa i filtri")

                    except Exception as e:
                        if args.debug:
                            print(
                                f"DEBUG: Errore elaborazione messaggio {msg_id}: {e}")
                        continue

                msg_ids = filtered_msg_ids
                print(
                    f"Messaggi che rispettano i filtri in {folder_name}: {len(msg_ids)}")

            if folder_matches_filters:
                folder_messages_dict[folder_name] = msg_ids
                total_messages_found += len(msg_ids)

        except Exception as e:
            print(
                f"Errore nell'elaborazione della cartella {folder_name}: {e}")
            continue

    # NUOVO: Controlla se ci sono cartelle che soddisfano i filtri
    print(f'\nTotale messaggi trovati: {total_messages_found}')
    print(f'Cartelle che soddisfano i filtri: {len(folders_matching_filters)}')

    if len(folders_matching_filters) == 0:
        print('Nessuna cartella soddisfa i criteri di ricerca.')
        imap.logout()
        sys.exit(0)

    # MOSTRA RIEPILOGO MESSAGGI PER CARTELLA
    print("\nRiepilogo cartelle e messaggi:")
    for folder_name in folders_matching_filters:
        msg_count = len(folder_messages_dict.get(folder_name, []))
        print(f"  - {folder_name}: {msg_count} messaggi")

    # GESTIONE OPZIONI MAILDIR
    if args.copy_to_maildir or args.move_to_maildir:
        maildir_path = args.copy_to_maildir or args.move_to_maildir
        move_mode = bool(args.move_to_maildir)

        if not create_maildir_structure(maildir_path, args.debug):
            print("Errore nella creazione della struttura MAILDIR")
            imap.logout()
            sys.exit(1)

        action = "spostamento" if move_mode else "copia"
        total_folders = len(folders_matching_filters)

        print(
            f"\nLa {action} creerà strutture MAILDIR per {total_folders} cartelle:")
        for folder_name in folders_matching_filters:
            msg_count = len(folder_messages_dict.get(folder_name, []))
            print(f"  - {folder_name} → {msg_count} messaggi")

        confirm_action = input(
            f"\nVuoi procedere con la {action} di {total_folders} cartelle (totale {total_messages_found} messaggi)? (s/n): ")

        if confirm_action.lower() != 's':
            print('Operazione annullata.')
            imap.logout()
            sys.exit(0)

        successful, failed = process_messages_to_maildir_multi_folder(
            imap, folder_messages_dict, maildir_path, base_folder,
            move_mode, args.expunge, args.debug
        )

        print(
            f"Operazione MAILDIR completata: {successful} messaggi elaborati con successo")
        if failed > 0:
            print(
                f"Attenzione: {failed} messaggi non sono stati elaborati correttamente")

        # NUOVO: Mostra struttura creata
        print(f"\nStruttura MAILDIR creata in: {maildir_path}")
        for folder_name in folders_matching_filters:
            relative_path = folder_name[len(base_folder):].lstrip('/')
            if relative_path:
                folder_path = os.path.join(
                    maildir_path, relative_path.replace('/', '.'))
            else:
                folder_path = maildir_path
            print(f"  - {folder_path}")

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
