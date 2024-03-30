import csv
import email
import imaplib
import re
from dataclasses import dataclass
from typing import List, Dict
import msal
from office365.graph_client import GraphClient
from office365.runtime.auth.token_response import TokenResponse


@dataclass
class Pyemail:
    host: str
    port: int
    user: str
    password: str
    mailboxes: List[str]
    provider: str
    secret_id: str = None
    secret_value: str = None
    client_id: str = None
    tenant_id: str = None

    # imap: imaplib.IMAP4_SSL = None

    def parse_mail_list(self, input):
        pattern = r'"([^"]*)"|([^"]+)'
        result = re.findall(pattern, input.decode('utf-8'))

        return result

    def get_inbox(self, mailbox):
        print(self.user, self.password)
        with imaplib.IMAP4_SSL(host=self.host, port=self.port) as imap:
            imap.login(self.user, password=self.password)

            # mail_list = imap.list()
            # for i in mail_list:
            #     for item in i[2:]:
            #         print(item)
            #         box = self.parse_mail_list(item)
            #         print(box)

            imap.select(mailbox=mailbox)
            _, msgnums = imap.search(None, 'ALL')
            rows = []
            for msgnum in msgnums[0].split():
                try:
                    _, data = imap.fetch(msgnum, "RFC822")
                    message = email.message_from_bytes(data[0][1])
                    message_from = message.get("FROM")
                    message_to = message.get("TO")
                    try:
                        contact_from, email_contact = message_from.split(" <")
                        email_from = email_contact[:-1]
                    except:
                        contact_from = ''
                        email_from = message_from
                    row = (contact_from, email_from, self.provider, self.user)
                    print(row)
                    rows.append(row)
                except Exception as e:
                    print(e)
                    continue

        # self.imap.close()
        # self.imap.logout()

        return rows


    def get_sent(self, mailbox):
        with imaplib.IMAP4_SSL(host=self.host, port=self.port) as imap:
            imap.login(self.user, password=self.password)
            imap.select(mailbox=mailbox)
            _, msgnums = imap.search(None, 'ALL')
            rows = []
            for msgnum in msgnums[0].split():
                try:
                    _, data = imap.fetch(msgnum, "RFC822")
                    message = email.message_from_bytes(data[0][1])
                    to_contact = message.get("TO")
                    print(to_contact)
                    try:
                        name, email_contact = to_contact.split(" <")
                        email_contact = email_contact[:-1]
                    except:
                        name = ''
                        email_contact = to_contact
                    row = (name, email_contact)
                    rows.append(row)
                except Exception as e:
                    print(e)
                    continue

        # self.imap.close()
        # self.imap.logout()

        return rows

    def acquire_token(self):
        """
        Acquire token via MSAL
        """
        authority_url = f'https://login.microsoftonline.com/{self.tenant_id}'
        app = msal.ConfidentialClientApplication(
            authority=authority_url,
            client_id=self.client_id,
            client_credential=self.secret_value,
        )
        token = app.acquire_token_for_client(scopes=[
            # "https://graph.microsoft.com/Mail.Read",
            # "https://graph.microsoft.com/Mail.ReadBasic",
            # "https://graph.microsoft.com/Mail.ReadBasic.All",
            # "https://graph.microsoft.com/Mail.ReadWrite",
            # "https://graph.microsoft.com/Mail.Send",
            # "https://graph.microsoft.com/MailboxSettings.Read",
            # "https://graph.microsoft.com/MailboxSettings.ReadWrite",
            # "https://graph.microsoft.com/User.Read",
            # "https://graph.microsoft.com/.default"
            # "offline_access Mail.ReadWrite Mail.send"
            "Mail.Read",
            "Mail.ReadBasic",
            "Mail.ReadBasic.All",
            "Mail.ReadWrite",
            "Mail.Send",
            "MailboxSettings.Read",
            "MailboxSettings.ReadWrite",
            "User.Read",
            ".default"
            # "offline_access Mail.ReadWrite Mail.send"
        ]
        )
        # result = TokenResponse(**token)
        # print(result)
        return token

    def get_o365_inbox(self):
        client = GraphClient(self.acquire_token)
        mail = client.me.mail_folders.execute_query()
        print(mail)

    def main(self):
        if self.provider != 'office_365':
            rows = []
            print(self.host, self.port, self.user, self.password, self.mailboxes)
            for mailbox in self.mailboxes:
                if mailbox == 'INBOX':
                    rows.extend(self.get_inbox(mailbox))
                # else:
                #     rows.extend(self.get_sent(mailbox))

                    # print(f'Message Number: {msgnum}')
                    # print(f'From: {message.get("From")}')
                    # print(f'To: {message.get("To")}')
                    # print(f'BCC: {message.get("BCC")}')
                    # print(f'Date: {message.get("Date")}')
                    # print(f'Subject: {message.get("Subject")}')

                    # print(f'Content:')
                    # for part in message.walk():
                    #     if part.get_content_type() == "text/plain":
                    #         print(part.as_string())

            result = set(rows)

            # self.imap.logout()

            with open('test.csv', 'a', newline='\n') as out:
                csv_out = csv.writer(out)
                csv_out.writerow(['name', 'email', 'provider', 'user'])
                for row in result:
                    csv_out.writerow(row)
        else:
            rows = []
            self.acquire_token()
            self.get_o365_inbox()


if __name__ == '__main__':
    with open('details.csv', newline='\n') as f:
        reader = csv.reader(f)
        email_lists = list(reader)[1:]

    configs = [
        {'provider': '1and1',
         'host': 'imap.1and1.com',
         'port': 993,
         'mailboxes': ['INBOX', '"Sent Items"']
         },
        {'provider': 'office_365',
         'host': 'outlook.office365.com',
         'port': 993,
         'mailboxes': ['INBOX', '"Sent Items"']
         },
        {'provider': 'cpanel',
         'host': '',
         'port': 993,
         'mailboxes': ['INBOX', '"Sent Items"']
         },
        {'provider': 'godaddy',
         'host': 'imap.secureserver.net',
         'port': 993,
         'mailboxes': ['INBOX', '"Sent Items"']
         },
    ]

    for an_email in email_lists:
        if an_email[0] != 'office_365':
            for config in configs:
                if an_email[0] == config['provider']:
                    used_config = config
                    break
                else:
                    continue
            if used_config['provider'] == 'cpanel':
                used_config['host'] = f'mail.{an_email[1].split("@")[1]}'

            user = an_email[1]
            password = an_email[2]

            m = Pyemail(host=used_config['host'], port=used_config['port'], user=user, password=password,
                        mailboxes=used_config['mailboxes'], provider=used_config['provider'])
            m.main()

        else:
            for config in configs:
                if an_email[0] == config['provider']:
                    used_config = config
                    break
                else:
                    continue

            user = an_email[1]
            password = an_email[2]
            secret_id = an_email[3]
            secret_value = an_email[4]
            client_id = an_email[5]
            tenant_id = an_email[6]
            m = Pyemail(host=used_config['host'], port=used_config['port'], user=user, password=password,
                        mailboxes=used_config['mailboxes'], provider=used_config['provider'], secret_id=secret_id,
                        secret_value=secret_value, client_id=client_id, tenant_id=tenant_id)
            m.main()