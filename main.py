import argparse
import logging
import time
from datetime import datetime, timezone
import json
import os
import re
import sys
from typing import Union, Optional

import mechanize
import pydantic
from imap_tools import MailBox, AND, MailboxLoginError
from mechanize import FormNotFoundError, Request

EUSERV_BASE = 'https://support.euserv.com/'
EUSERV_BASE_INDEX = 'https://support.euserv.com/index.iphp'
PATH_COOKIES = os.path.join('cookies.txt')
PATH_CONFIG = os.path.join('config.json')

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.WARNING)
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.DEBUG)


def isLoggedInEuserv(html: Union[str, None]) -> bool:
    if html is None:
        return False
    elif 'action=logout' in html:
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
    last_date_extended_contract: Optional[Union[datetime, None]]
    last_date_login_attempt_failed: Optional[Union[datetime, None]]
    last_date_login_attempt_failed_captcha: Optional[Union[datetime, None]]
    last_contract_id: Optional[str]
    last_phpsessid: Optional[str]
    last_customer_id: Optional[str]


def getConfig() -> Config:
    configpath = os.path.join(PATH_CONFIG)
    print(f'Loading config from {configpath}')
    with open(configpath, encoding='utf-8') as infile:
        jsondict = json.load(infile)
        return Config(**jsondict)


class ContractUpdater:

    def __init__(self):
        try:
            self.config = getConfig()
        except Exception as e:
            print(e)
            print('Kaputte config: Ein- oder mehrere Eintraege fehlen!')
            sys.exit()
        self.br = getNewBrowser()
        self.mailbox = MailBox(self.config.imap_server)

    def getMainpageURL(self, sessionid: str) -> str:
        return EUSERV_BASE_INDEX + '?sess_id=' + sessionid + '&action=show_default'

    def ensureMailLogin(self):
        if self.mailbox.login_result is None:
            """ We're not yet logged in -> Login """
            try:
                self.mailbox.login(self.config.imap_login, self.config.imap_password)
            except MailboxLoginError as me:
                print("Ungueltige Email Zugangsdaten")
                raise me

    def loginEuserv(self):
        # TODO: Fix loadCookies handling (?)
        last_phpsessid = self.config.last_phpsessid
        response = None
        html = None
        if len(self.br.cookiejar) > 0 and last_phpsessid is not None:
            # Try to login via stored cookies first
            print(f'Versuche Login ueber zuvor gespeicherte Cookies mit {last_phpsessid=} ...')
            # self.br.addheaders.append(('Referer', EUSERV_BASE_INDEX + '?sess_id=' + last_phpsessid))
            response = self.br.open(self.getMainpageURL(last_phpsessid))
            html = getHTML(response)
            if isLoggedInEuserv(html):
                print("Cookie login erfolgreich")
                self.br.cookiejar.save()
                return
            print("Cookie login fehlgeschlagen")
        print("Versuche vollständigen Login")
        self.br.cookiejar.clear()
        response = self.br.open(EUSERV_BASE_INDEX)
        html = getHTML(response)
        sess_id_regex = re.compile(r'sess_id=([a-f0-9]+)')
        sess_idRegex = sess_id_regex.search(html)
        # if sess_idRegex is None:
        #     # Obtain ID from current URL
        #     sess_idRegex = sess_id_regex.search(self.br.geturl())
        if sess_idRegex is None:
            raise Exception("Konnte session_id nicht finden")
        sess_id = sess_idRegex.group(1)
        print("Aktive Session: " + sess_id)
        # Important!
        self.br.open("https://support.euserv.com/pic/logo_small.png")
        self.br.open(EUSERV_BASE_INDEX + '?sess_id=' + sess_id)
        html = getHTML(response)
        try:
            self.br.select_form(name_=lambda x: 'step1_anmeldung' in x)
            self.br.form.set_all_readonly(False)
            # Allow us to change those fields (stupid workaround)
            self.br.form.new_control('text', 'email', {'email': ''})
            self.br.form.new_control('text', 'password', {'password': ''})
            self.br.form.new_control('text', 'form_selected_language', {'form_selected_language': ''})
            self.br.form.new_control('text', 'Submit', {'Submit': ''})
            self.br.form.new_control('text', 'subaction', {'subaction': ''})
            self.br.form.new_control('text', 'sess_id', {'sess_id': ''})

            self.br.form.fixup()
            self.br['email'] = self.config.euserv_mail_or_user_id
            self.br['password'] = self.config.euserv_password
            self.br['form_selected_language'] = 'de'
            self.br['Submit'] = 'Anmelden'
            self.br['subaction'] = 'login'
            self.br['sess_id'] = sess_id
            self.br.set_handle_referer(False)
            self.br.addheaders.append(('Referer', self.br.geturl()))
            response = self.br.submit()
            self.br.set_handle_referer(True)
            html = getHTML(response)
        except FormNotFoundError:
            print('Loginfehler: Konnte Loginform nicht finden')
            # Do not give up yet - maybe only PIN-step is needed
        try:
            self.br.select_form(name_=lambda x: 'step2_anmeldung' in x)
            print('PIN benoetigt')
            customer_id_regex = re.compile(r"name=\"c_id\"[^>]*value=\"(\d+)\"").search(html)
            if customer_id_regex is None:
                raise Exception("Konnte c_id nicht finden")
            customer_id = customer_id_regex.group(1)
            pin = self.mailFindLoginPIN()
            self.br.form.set_all_readonly(False)
            # Allow us to change those fields (stupid workaround)
            self.br.form.new_control('text', 'pin', {'pin': ''})
            self.br.form.new_control('text', 'save_for_auto_login', {'save_for_auto_login': ''})
            self.br.form.new_control('text', 'Submit', {'Submit': ''})
            self.br.form.new_control('text', 'subaction', {'login': ''})
            self.br.form.new_control('text', 'sess_id', {'sess_id': ''})
            self.br.form.new_control('text', 'c_id', {'c_id': ''})
            self.br.form.fixup()
            self.br['pin'] = pin
            self.br['save_for_auto_login'] = 'on'
            self.br['Submit'] = 'Besttigen'
            self.br['subaction'] = 'login'
            self.br['sess_id'] = sess_id
            self.br['c_id'] = customer_id
            self.config.last_customer_id = customer_id
            response = self.br.submit()
            html = getHTML(response)
        except FormNotFoundError:
            print('Keine PIN benoetigt')
        if not isLoggedInEuserv(html):
            if 'securimage_show' in html:
                print("Euserv Login fehlgeschlagen - Login-Captcha benötigt!")
                self.config.last_date_login_attempt_failed_captcha = datetime.now().now().timestamp()
            else:
                print('Euserv Login fehlgeschlagen - Ungueltige Zugangsdaten?')
                self.config.last_date_login_attempt_failed = datetime.now().now().timestamp()
            print('html: ' + html)
            self.saveConfig()
            raise Exception("Login fehlgeschlagen")
        phpsessid = self.browserGetCookieByKey('PHPSESSID')
        if phpsessid is None:
            # This should never happen!
            raise Exception("Ungueltiger Loginstatus")
        print('Euserv Login erfolgreich')
        self.config.last_phpsessid = sess_id
        # TODO: Remove those fields on successful login
        # self.config.last_date_login_attempt_failed = None
        # self.config.last_date_login_attempt_failed_captcha = None
        # Store cookies in file so we can try to re-use them next time
        self.br.cookiejar.save(ignore_expires=True)
        self.saveConfig()



    def browserGetCookieByKey(self, key: str) -> Union[str, None]:
        for cookie in self.br.cookiejar:
            if cookie.name == key:
                return cookie.value
        return None

    def mailFindContractID(self) -> str:
        """ Sucht nach "Vertragsverlängerungs Emails" und gibt die erste ID eines verlängerbaren Vertrags zurück. """
        self.ensureMailLogin()
        print('Sammle Vertragsverlaengerungs-E-Mails ...')
        lastHoursToCheck = 48
        timestamp = datetime.now().timestamp() - lastHoursToCheck * 60 * 60
        targetdate = datetime.fromtimestamp(timestamp).date()
        """ Solange wie die ID des Vertrags im Betreff steht, können wir uns das Laden vom Inhalt der Mail sparen :) """
        headersOnly = True
        mails = self.mailbox.fetch(mark_seen=False, criteria=AND(subject='Anstehende manuelle Vertragsverlaengerung fuer Vertrag', date_gte=targetdate), headers_only=headersOnly)
        contractIDs = []
        for msg in mails:
            emailText = msg.text
            if headersOnly:
                regex = re.compile(r'(?i)fuer Vertrag (\d+)').search(msg.subject)
                if regex is None:
                    raise Exception("Fatal: Failed to find contractID in subject of email:\n" + emailText)
                contractID = regex.group(1)
            else:
                regex = re.compile(r'(?i)Sie den Vertrag (\d+) aus').search(emailText)
                if regex is None:
                    raise Exception("Fatal: Failed to find contractID in text of email:\n" + emailText)
                contractID = regex.group(1)
            contractIDs.append(contractID)
        if len(contractIDs) == 0:
            raise Exception(f"Keine Vertragsverlaengerungsemail in Mails der letzten {lastHoursToCheck} Stunden gefunden")
        contractIDToUse = contractIDs[0]
        if len(contractIDs) > 1:
            print(f"Es wurden mehrere verlängerbare VertragsIDs gefunden: {contractIDs} -> Script verwendet ID {contractIDToUse}")
        return contractIDToUse

    def mailFindLoginPIN(self) -> str:
        self.ensureMailLogin()
        secondsWaited = 0
        secondsWaitMax = 600
        secondsWaitPerLoop = 10
        print("Warte auf Login-PIN...")
        while secondsWaited < secondsWaitMax:
            time.sleep(secondsWaitPerLoop)
            secondsWaited += secondsWaitPerLoop
            timestamp = datetime.now().timestamp() - 5 * 60
            targetdate = datetime.fromtimestamp(timestamp).date()
            mails = self.mailbox.fetch(mark_seen=False, criteria=AND(subject='EUserv - Versuchter Login', date_gte=targetdate), reverse=True)
            for msg in mails:
                emailBody = msg.html
                regex = re.compile(r'(?i)PIN\s*:\s*<br>\s*(\d{6})').search(emailBody)
                if regex is None:
                    # Maybe forwarded email?
                    regex = re.compile(r'(?i)PIN\s*:\s*\r\n(\d{6})').search(msg.text)
                    if regex is None:
                        raise Exception("Fatal: Failed to find PIN in PIN email:\n" + msg.html + " | Text: " + msg.text)
                loginpin = regex.group(1)
                print("Login PIN gefunden: " + loginpin)
                return loginpin
            print(f"Noch keine PIN-Mail gefunden | Sekunden gewartet: {secondsWaited}/{secondsWaitMax}")
        raise Exception("Fata: PIN-Mail nicht gefunden")

    def mailFindContractExtendPIN(self) -> str:
        self.ensureMailLogin()
        secondsWaited = 0
        secondsWaitMax = 600
        secondsWaitPerLoop = 10
        print(f"Warte auf Vertragsverlängerungs-PIN max {secondsWaitMax} Sekunden mit {secondsWaitPerLoop} Sekunden Abstand pro Versuch...")
        pins = set()
        firstPIN = None
        while secondsWaited < secondsWaitMax:
            time.sleep(secondsWaitPerLoop)
            secondsWaited += secondsWaitPerLoop
            print(f"Warte auf Vertragsverlängerungs-PIN | Sekunden gewartet: {secondsWaited}/{secondsWaitMax}")
            timestamp = datetime.now().timestamp() - 5 * 60
            targetdate = datetime.fromtimestamp(timestamp).date()
            mails = self.mailbox.fetch(mark_seen=False, criteria=AND(subject='EUserv - PIN zur', date_gte=targetdate), reverse=True)
            for msg in mails:
                emailBody = msg.html
                regex = re.compile(r'(?i)PIN\s*:\s*<br>\s*(\d{6})').search(emailBody)
                if regex is None:
                    # Maybe forwarded email?
                    regex = re.compile(r'(?i)PIN\s*:\s*\r\n(\d{6})').search(msg.text)
                    if regex is None:
                        raise Exception("Fatal: Konnte Vertragsverlängerungspin nicht finden in Email:\n" + msg.html + " | Text: " + msg.text)
                pin = regex.group(1)
                pins.add(pin)
                if firstPIN is None:
                    firstPIN = pin
            if len(pins) > 0:
                break
        if len(pins) == 0:
            raise Exception("Fata: VertragsverlängerungsPIN-Mail nicht gefunden")
        if len(pins) > 1:
            print(f"Warnung: Es wurden mehrere ({len(pins)}) Vertragsverlängerungs-PINs gefunden: {pins} | Es wird nur die erste probiert")
        return firstPIN

    def extendContract(self, contractID: str):
        self.loginEuserv()
        mainurl = self.getMainpageURL(self.config.last_phpsessid)
        print(f"Vorbereitung: Öffne Startseite: {mainurl}")
        html = getHTML(self.br.open(mainurl))
        contractExtendUrlRegex = re.compile(r"\"(/index\.iphp\?[^\"]*show_contract_extension=1[^\"].*" + contractID + "[^\"]*)").search(html)
        if contractExtendUrlRegex is None:
            raise Exception("Fatal: Failed to find prolongContractURL")
        phpsessid = self.browserGetCookieByKey('PHPSESSID')
        prolongContractURL = contractExtendUrlRegex.group(1)
        print("Starte Vertragsverlängerung")
        numberofSteps = 6
        print(f"Schritt 1/{numberofSteps}: Öffne URL: {prolongContractURL}")
        self.br.open(prolongContractURL)
        data = {
            'sess_id': phpsessid,
            'subaction': 'show_kc2_security_password_dialog',
            'prefix': 'kc2_customer_contract_details_extend_contract_',
            'type': 1
        }
        print("Schritt 2: Frage Vertragsverängerung an")
        self.br.open(Request(url=EUSERV_BASE_INDEX, data=data))
        print("Schritt 3: Warte auf Email mit PIN")
        contractExtendPIN = self.mailFindContractExtendPIN()
        data['auth'] = contractExtendPIN
        data['subaction'] = 'kc2_security_password_get_token'
        data['ident'] = 'kc2_customer_contract_details_extend_contract_' + contractID
        print("Schritt 4: Sende PIN " + contractExtendPIN)
        response = self.br.open(Request(url=EUSERV_BASE_INDEX, data=data))
        parsedJson = json.loads(getHTML(response))
        status = parsedJson.get('rs')
        if status != 'success':
            # This should never happen!
            print("!FATAL: Vertragsverlängerung fehlgeschlagen! | Fehler: " + status)
            sys.exit()
        token = parsedJson['token']['value']
        data2 = {
            'sess_id': phpsessid,
            'subaction': 'kc2_customer_contract_details_get_extend_contract_confirmation_dialog',
            'token': token
        }
        print("Schritt 5: Bestätige Vertragsverlängerung Teil 1")
        response = self.br.open(Request(url=EUSERV_BASE_INDEX, data=data2))
        jsonText = getHTML(response)
        parsedJson = json.loads(jsonText)
        html = parsedJson['html']['value']
        contractExtendDateRegex = re.compile(r'bis zum (\d{2}\.\d{2}\.\d{4})').search(html)
        contractExtendDate = None
        if contractExtendDateRegex is not None:
            contractExtendDate = contractExtendDateRegex.group(1)
            self.config.last_date_extended_contract = contractExtendDate
        else:
            print("Warnung: Verlängerungsdatum konnte nicht gefunden werden | HTML:")
            print(html)
        print(f"Schritt 6: Bestätige Vertragsverlängerung Teil 2: Vertrag {contractID} wird verlängert bis: {contractExtendDate}")
        data3 = {
            'sess_id': phpsessid,
            'ord_id': contractID,
            'subaction': 'kc2_customer_contract_details_extend_contract_term',
            'token': token
        }
        response = self.br.open(Request(url=EUSERV_BASE_INDEX, data=data3))
        html = getHTML(response)
        if 'Der Vertrag wurde verl' in html or (contractExtendDate is not None and contractExtendDate in html):
            print(f"Vertrag {contractID} wurde erfolgreich verlängert | Die Verlängerung wird dir zusätzlich per Email bestätigt.")
        else:
            print('Warnung: Vertragsverlängerung evtl. fehlgeschlagen -> Prüfe deine Emails!')
        self.config.last_date_extended_contract = datetime.now().now().timestamp()
        # print("Finale Antwort: " + html)

    def saveConfig(self):
        saveJson(jsonData=self.config.dict(), filepath=PATH_CONFIG)

    def run(self):
        lastTimeContraxtExtended = self.config.last_date_extended_contract
        minDaysUntilNextAttempt = 3
        debug = False
        if lastTimeContraxtExtended is not None and (datetime.now(tz=timezone.utc) - lastTimeContraxtExtended).days <= minDaysUntilNextAttempt and not debug:
            # Check if contract was extended within the last 48 hours -> Then don't even try
            print(f"Dein Vertrag wurde zuletzt per Script verlängert vor weniger als {minDaysUntilNextAttempt} Tagen am {lastTimeContraxtExtended} -> Beende ohne Aktion")
            sys.exit()
        contractID = self.mailFindContractID()
        self.config.last_contract_id = contractID
        print(f'Vertrags-ID zum Verlaengern gefunden: {contractID} | Letzte Vertrags-ID aus Config: {self.config.last_contract_id}')
        self.extendContract(contractID)

        print("Speichere Config und beende")
        self.saveConfig()
        sys.exit()


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
    cj = mechanize.LWPCookieJar(PATH_COOKIES)
    br.set_cookiejar(cj)
    if os.path.exists(PATH_COOKIES):
        br.cookiejar.load()

    br.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'),
                     ('Origin', 'https://support.euserv.com'),
                     ('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7'),
                     ('Accept-Language', 'de-DE,de;q=0.9,en;q=0.8,en-US;q=0.7'),
                     ('Cache-Control', 'max-age=0'),
                     ('Accept-Encoding', 'gzip, deflate, br'),
                     ('Upgrade-Insecure-Requests', '1'),
                     ('Sec-Ch-Ua', "\"Not/A)Brand\";v=\"99\", \"Google Chrome\";v=\"115\", \"Chromium\";v=\"115\""),
                     ('Sec-Ch-Ua-Mobile', '?0'),
                     ('Sec-Ch-Ua-Platform', "\"Windows\""),
                     ('Sec-Fetch-Dest', 'document'),
                     ('Sec-Fetch-Mode', 'navigate'),
                     ('Sec-Fetch-Site', 'same-origin'),
                     ('Sec-Fetch-User', '?1'),
                     ('Connection', 'keep-alive')
                     ]
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

    euservContractUpdater = ContractUpdater()
    if args.test_logins:
        euservContractUpdater.ensureMailLogin()
        print("Email Login OK")
        euservContractUpdater.loginEuserv()
        print("Euserv Login OK")
        sys.exit()
    else:
        euservContractUpdater.run()
