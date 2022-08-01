import imaplib
import json
import re
import sys


class ContractUpdater:

    def __init__(self, cfg: dict):
        # TODO: Add errorhandling for missing keys
        self.imap_server = cfg['imap_server']
        self.login_imap = cfg['login_imap']
        self.password_imap = cfg['password_imap']



    def run(self):
        try:
            connection = imaplib.IMAP4_SSL(self.imap_server)
            connection.login(self.login_imap, self.password_imap)
            print('E-Mail Login erfolgreich')
        except Exception as e:
            print(e)
            print('E-Mail Login fehlgeschlagen!')
            print('Falls du GMail Benutzer bist aktiviere den Zugriff durch weniger sichere Apps: https://myaccount.google.com/lesssecureapps')
            sys.exit()
        typ, data = connection.list()
        if typ != 'OK':
            # E.g. NO = Invalid mailbox (should never happen)
            print("Keine Postfaecher gefunden")
            return
        # print('Anzahl gefundener Postfaecher: %d' % len(data))
        # this maybe used to speed-up the checking process - users can blacklist postboxes through this list
        postbox_ignore = ['Sent']
        numberof_postboxes = len(data)
        postbox_index = 0
        total_postbox_steps = 2
        mails = []
        for line in data:
            postbox_index += 1
            flags, delimiter, mailbox_name = parse_list_response(line)
            print('Arbeite an Postfach %d / %d: \'%s\' ...' % (postbox_index, numberof_postboxes, mailbox_name))
            if mailbox_name in postbox_ignore:
                # Rare case
                print(f'Ueberspringe aktuelles Postfach {mailbox_name}, da es sich auf der Blacklist befindet')
                continue
            # Surround mailbox_name  with brackets otherwise this will fail for email labels containing spaces
            typ, data = connection.select('"' + mailbox_name + '"', readonly=True)
            if typ != 'OK':
                # E.g. NO = Invalid mailbox (should never happen)
                print(f'Fehler: Postfach {mailbox_name} konnte nicht geoeffnet werden')
                # 2020-01-03: Skip invalid mailboxes - this may happen frequently with gmail accounts (missing permissions?)
                continue
            # Search for specific messages by subject
            print(f'Schritt 1 / {total_postbox_steps}: Sammle Vertragsverlaengerungs-E-Mails ...')
            mails += crawlMailsBySubject(connection, 'Anstehende manuelle Vertragsverlaengerung fuer Vertrag')
        serverContractID = None
        for mail in mails:
            regex = re.compile(r'Vertrag\s*[^:]+:\s*(\d+)').search(mail)
            if regex is not None:
                serverContractID = regex.group(1)
                break
        if serverContractID is None:
            print("Es wurde keine Mail zur anstehenden Vertragsverlaengerung gefunden")
            sys.exit()
        print('Vertrags-ID zum Verlaengern gefunden: ' + serverContractID)


def crawlMailsBySubject(connection, subject: str) -> list:
    typ, [msg_ids] = connection.search(None, '(SUBJECT "' + subject + '")')
    # print('INBOX', typ, msg_ids)
    emails = []
    for msg_id in msg_ids.split():
        # print('Fetching mail with ID %s' % msg_id)
        typ, msg_data = connection.fetch(msg_id, '(BODY.PEEK[TEXT])')
        complete_mail = ''
        for response_part in msg_data:
            # print('Printing response part:')
            if isinstance(response_part, tuple):
                # print('\n%s:' % msg_id)
                mail_part = response_part[1].decode('utf-8', 'ignore')
                # print(mail_part)
                complete_mail += mail_part
        emails.append(complete_mail)
    return emails

def loadJson(filepath: str):
    readFile = open(filepath, 'r')
    settingsJson = readFile.read()
    readFile.close()
    return json.loads(settingsJson)


list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')
# According to docs/example: https://pymotw.com/2/imaplib/
def parse_list_response(line):
    line = line.decode('utf-8', 'ignore')
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return flags, delimiter, mailbox_name

if __name__ == '__main__':
    # TODO: Add errorhandling
    config = loadJson('config.json')
    euservContractUpdater = ContractUpdater(cfg=config)
    euservContractUpdater.run()

