import argparse
from datetime import datetime
import imaplib
import json
import os
import re
import sys

import mechanize
from mechanize import FormNotFoundError

EUSERV_BASE = 'https://support.euserv.com/'
PATH_COOKIES = os.path.join('cookies.txt')
PATH_SAVESTATE = os.path.join('savestate.json')

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
        except Exception as ex:
            print(ex)
            print('E-Mail Login fehlgeschlagen!')
            print('Falls du GMail Benutzer bist, aktiviere den Zugriff durch weniger sichere Apps hier: https://myaccount.google.com/lesssecureapps')
            return None

    def loginEuserv(self):
        br = getNewBrowser()
        # TODO: Fix loadCookies handling
        cookies = mechanize.LWPCookieJar(PATH_COOKIES)
        if cookies is not None and os.path.exists(PATH_COOKIES):
            # Try to login via stored cookies first
            print('Versuche Login ueber zuvor gespeicherte Cookies ...')
            br.set_cookiejar(cookies)
        response = br.open(EUSERV_BASE)
        html = getHTML(response)
        if not isLoggedInEuserv(html):
            if cookies is not None and os.path.exists(PATH_COOKIES):
                print('Login ueber Cookies fehlgeschlagen --> Versuche vollstaendigen Login')
            try:
                br.select_form(name_=lambda x: 'step1_anmeldung' in x)
                # WTF form is found but all fields are missing
                sess_idRegex = re.compile(r'sess_id=([a-f0-9]+)').search(html)
                sess_id = None
                if sess_idRegex:
                    sess_id = sess_idRegex.group(1)
                # br.form.set_all_readonly(False)
                br.form.new_control('text', 'email', {'email': ''})
                br.form.new_control('text', 'password', {'password': ''})
                br.form.new_control('text', 'form_selected_language', {'form_selected_language': ''})
                br.form.new_control('text', 'Submit', {'Submit': ''})
                br.form.new_control('text', 'subaction', {'subaction': ''})
                br.form.new_control('text', 'sess_id', {'sess_id': ''})

                br.form.fixup()
                br['email'] = self.euserv_login
                br['password'] = self.euserv_password
                br['form_selected_language'] = 'de'
                br['Submit'] = 'Anmelden'
                br['subaction'] = 'login'
                br['sess_id'] = sess_id
                response = br.submit()
                html = getHTML(response)
            except FormNotFoundError:
                print('Konnte Loginform nicht finden - evtl. wird nur 2FA login benoetigt...')
                sys.exit()
        # TODO: Simplify this
        if setFormBySubmitKey(br, 'pin'):
            # TODO: Add support for 2FA login
            pass
        if not isLoggedInEuserv(html):
            print('Login fehlgeschlagen - Ungueltige Zugangsdaten?')
            return br, False
        print('Euserv Login erfolgreich')
        # Store cookies so we can re-use them next time
        cookies = br._ua_handlers['_cookies'].cookiejar
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
        print('Sammle Vertragsverlaengerungs-E-Mails ...')

        # This can be used to speed-up the checking process - users can blacklist postboxes through this list
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
        br, loggedIn = self.loginEuserv()
        if not loggedIn:
            sys.exit()
        saveJson(jsonData={'': datetime.now().timestamp}, filepath=PATH_SAVESTATE)
        return None

def setFormBySubmitKey(br, submitKey: str) -> bool:
    if submitKey is None:
        return False
    current_index = 0
    for form in br.forms():
        for control in form.controls:
            if control.name is None:
                continue
            if control.name == submitKey:
                br.select_form(nr=current_index)
                return True
        current_index += 1
    return False


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

def saveJson(jsonData, filepath):
    with open(filepath, 'w') as outfile:
        json.dump(jsonData, outfile)

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
    my_parser = argparse.ArgumentParser()
    my_parser.add_argument('-t', '--test_logins', help='Nur Logins testen und dann beenden.', type=bool, default=False)
    args = my_parser.parse_args()
    try:
        config = loadJson('config.json')
    except Exception as e:
        print("config.json fehlt!")
        sys.exit()

    euservContractUpdater = ContractUpdater(cfg=config)
    if args.test_logins:
        euservContractUpdater.loginMail()
        euservContractUpdater.loginEuserv()
        sys.exit()
    else:
        euservContractUpdater.run()
