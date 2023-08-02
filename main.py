import argparse
import logging
from datetime import datetime
import imaplib
import json
import os
import re
import sys
from pathlib import Path

import mechanize
import pydantic
from imap_tools import MailBox, AND, MailboxLoginError
from mechanize import FormNotFoundError

EUSERV_BASE = 'https://support.euserv.com/'
PATH_COOKIES = os.path.join('cookies.txt')
PATH_SAVESTATE = os.path.join('savestate.json')

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.WARNING)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.DEBUG)


def isLoggedInEuserv(html: str) -> bool:
    if 'action=logout' in html:
        return True
    else:
        return False


# Converts html bytes from response object to String
def getHTML(response):
    return response.read().decode('utf-8', 'ignore')


class Config(pydantic.BaseModel):
    imap_server: str
    imap_login: str
    imap_password: str
    euserv_mail_or_user_id: str
    euserv_password: str


def getConfig() -> Config:
    currentPath = Path(os.getcwd())
    # Go to parent as config is contained in root folder of project
    configpath = os.path.join('config.json')
    print(f'Loading config from {configpath}')
    with open(configpath, encoding='utf-8') as infile:
        jsondict = json.load(infile)
        return Config(**jsondict)


class ContractUpdater:

    def __init__(self, cfg: dict):
        try:
            self.config = getConfig()
        except Exception as e:
            print(e)
            print('Kaputte config: Ein- oder mehrere Eintraege fehlen!')
            sys.exit()

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
                br['email'] = self.config.euserv_mail_or_user_id
                br['password'] = self.config.euserv_password
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
            print('Euserv Login fehlgeschlagen - Ungueltige Zugangsdaten?')
            return br, False
        print('Euserv Login erfolgreich')
        # Store cookies so we can re-use them next time
        cookies = br._ua_handlers['_cookies'].cookiejar
        cookies.save()
        return br, True

    def run(self):
        print('Sammle Vertragsverlaengerungs-E-Mails ...')

        try:
            with MailBox(self.config.imap_server).login(self.config.imap_login, self.config.imap_password) as mailbox:
                timestamp = datetime.now().timestamp() - 24 * 60 * 60
                targetdate = datetime.fromtimestamp(timestamp).date()
                mails = mailbox.fetch(mark_seen=False, criteria=AND(subject='Anstehende manuelle Vertragsverlaengerung fuer Vertrag', date_gte=targetdate))
                contractIDs = []
                for msg in mails:
                    emailBody = msg.html
                    regex = re.compile(r'(?i)Sie den Vertrag (\d+) aus').search(emailBody)
                    if regex is None:
                        raise Exception("Fatal: Failed to find contractID in email:\n" + emailBody)
                    contractID = regex.group(1)
                    contractIDs.append(contractID)
                if len(contractIDs) == 0:
                    print("Keine Vertragsverlaengerungsemail gefunden")
                    sys.exit()
                elif len(contractIDs) > 1:
                    pass
        except MailboxLoginError:
            print("Ungueltige Email Zugangsdaten")
            sys.exit()
        serverContractID = contractIDs[0]
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
                      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36')]
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
