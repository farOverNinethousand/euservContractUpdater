import imaplib
import json
import os
import re
import sys

import mechanize

EUSERV_BASE = 'https://support.euserv.com/'

def isLoggedInEuserv(html: str) -> bool:
    if 'action=logout' in html:
        return True
    else:
        return False

# Converts html bytes from response object to String
def getHTML(response):
    return response.read().decode('utf-8', 'ignore')


class ContractUpdater:


    def __init__(self, cfg: dict):
        # TODO: Add errorhandling for missing keys
        try:
            self.imap_server = cfg['imap_server']
            self.imap_login = cfg['imap_login']
            self.imap_password = cfg['imap_password']
            self.euserv_login = cfg['euserv_mail_or_user_id']
            self.euserv_password = cfg['euserv_password']
        except KeyError as e:
            print(e)
            print('Kaputte config: Ein- oder mehrere Eintraege fehlen!')
            sys.exit()

    def loginMail(self):
        try:
            connection = imaplib.IMAP4_SSL(self.imap_server)
            connection.login(self.imap_login, self.imap_password)
            print('E-Mail Login erfolgreich')
            return connection
        except Exception as e:
            print(e)
            print('E-Mail Login fehlgeschlagen!')
            print('Falls du GMail Benutzer bist, aktiviere den Zugriff durch weniger sichere Apps hier: https://myaccount.google.com/lesssecureapps')
            return None

    def loginEuserv(self):
        br = getNewBrowser()
        # TODO: Fix loadCookies handling
        # cookies = mechanize.LWPCookieJar(getCookiesPath())
        cookies = None
        if cookies is not None and os.path.exists(getCookiesPath()):
            # Try to login via stored cookies first
            print('Versuche Login ueber zuvor gespeicherte Cookies ...')
            br.set_cookiejar(cookies)
        response = br.open(EUSERV_BASE)
        html = getHTML(response)
        if not isLoggedInEuserv(html):
            if cookies is not None and os.path.exists(getCookiesPath()):
                print('Login ueber Cookies fehlgeschlagen --> Versuche vollstaendigen Login')
            br.open(EUSERV_BASE)
            foundForm = False
            form_index = 0
            for form in br.forms():
                if 'step1_anmeldung' in form:
                    foundForm = True
                    br.select_form(nr=form_index)
                    break
                form_index += 1
            if not foundForm:
                print('Fataler Fehler: Konnte Loginform nicht finden...')
                sys.exit()
            br['email'] = self.euserv_login
            br['password'] = self.euserv_password
            response = br.submit()
            html = getHTML(response)
            # TODO: Add support for 2FA login
            if not isLoggedInEuserv(html):
                print('Login fehlgeschlagen - Ungueltige Zugangsdaten?')
                return br, False
            print('Vollstaendiger Login erfolgreich')
        cookies = br._ua_handlers['_cookies'].cookiejar
        print('Speichere Cookies in ' + 'TODO')
        cookies.save()
        return br, True

    def run(self):
        connection = self.loginMail()
        if connection is None:
            # Login failure
            sys.exit()
        typ, data = connection.list()
        if typ != 'OK':
            # E.g. 'NO' = Invalid mailbox (should never happen)
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
        # br, loggedIn = self.loginEuserv()
        if not loggedIn:
            sys.exit()
        return None




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

def getNewBrowser():
    # Prepare browser
    br = mechanize.Browser()
    # br.set_all_readonly(False)    # allow everything to be written to
    br.set_handle_robots(False)  # ignore robots
    br.set_handle_refresh(False)  # can sometimes hang without this
    br.set_handle_referer(True)
    br.set_handle_redirect(True)
    br.addheaders = [('User-agent',
                      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36')]
    return br


list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')


# According to docs/example: https://pymotw.com/2/imaplib/
def parse_list_response(line):
    line = line.decode('utf-8', 'ignore')
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return flags, delimiter, mailbox_name


if __name__ == '__main__':
    # TODO: Add errorhandling
    try:
        config = loadJson('config.json')
    except Exception as e:
        print("config.json fehlt!")
        sys.exit()

    euservContractUpdater = ContractUpdater(cfg=config)
    euservContractUpdater.run()
