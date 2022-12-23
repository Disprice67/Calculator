from lib_checkmail import *


class Attachmail:

    R_EMAIL: str = None
    R_SUBJECT: str = None

    CONNECTION = None
    ERRORS = None

    def __init__(self, mail_server: str, 
                 username: str, 
                 password: str
                 ):

        self.connection = imaplib.IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select('INBOX')


    def close_connection(self):
        """Close connection"""

        self.connection.close()


    def search_count_messages_dir(self, dir: str):
        """Search count messages"""

        typ, items = self.connection.search(None, dir)
        items = items[0].split()
        count = len(items)
        return count


    def parse_uid(self, data: str):
        """Parse UID"""

        match = re.search(r'\d+ \(UID (?P<uid>\d+)\)', data)
        return match.group('uid')

    def send_email(self, user: str, 
                   pwd: str, 
                   recipient: str, 
                   subject: str, 
                   body: str,
                   file = None,
                   isTls=True,
                   filename = None):

        """Send email."""
        msg = EmailMessage()
        msg['From'] = user
        msg['Subject'] = subject
        msg['To'] = recipient
        msg.set_content(body)

        if file is not None:

            with open(file, 'rb') as f:
                file_data = f.read()

            msg.add_attachment(file_data, 
                                maintype="application", 
                                subtype="xlsx",
                                filename=filename
                                )

        with smtplib.SMTP_SSL(config.MAIL_SERVER, 465) as smtp:
            smtp.login(user, pwd) 
            smtp.send_message(msg)

    def get_attachments(self, attachment_dir: str,
                        e):

        """Getattachment in messages"""

        encode = unicode.UnicodeReader()
        body_messages = None
        list_name = []
        filename = None
        flag = False
        typ, items = self.connection.search(None, 'ALL')
        try:
            items = items[0].split()
            emailid = items[0]
        except:
            return False

        resp, data = self.connection.fetch(emailid, "(RFC822)")
        u_resp, u_data = self.connection.fetch(emailid, "(UID)")

        msg_uid = self.parse_uid(data=str(u_data[0]))

        #getting the mail content
        email_body = data[0][1]
        mail = email.message_from_bytes(email_body)

        #reply item
        self.R_EMAIL = mail['From'].split()[-1]
        self.R_SUBJECT = encode.encoded(words=mail['Subject'])
        subject = encode.encoded(words=mail['Subject'])

        # is this part an attachment?

        for part in mail.walk():
            
            if part.get_filename() is None:
                continue

            filename = encode.encoded(words=part.get_filename())

            if '.jpg' in filename or '.png' in filename or '.gif' in filename:
                continue

            flag = True

            filename += '.xlsx'
            
            list_name.append((filename, part))

        if flag is False:

            body_messages = (f"""Письмо с темой: {self.R_SUBJECT}.
Не подходит для обработки,
скорее всего в письме нет нужного формата файла!:(
""")
               
            #mail attribute
            send_email = self.send_email(
                            user= config.USERNAME_GMAIL,
                            pwd= config.PASSWORD_GMAIL,
                            recipient= self.R_EMAIL,
                            subject= self.R_SUBJECT,
                            body= body_messages
                            )

            mov, data = self.connection.uid('STORE', msg_uid, 
                                            '+FLAGS', '(\Deleted)')

            return False

        for file in list_name:
            
            file_name, part_t = file

            if subject == 'Обновление шаблона':
                
                if len(list_name) > 1:

                    body_messages = (f"""Письмо с темой: {self.R_SUBJECT}.
В письме для обновления шаблона больше одного файла, 
пришлите один файл в формате '.xlsx'.
""")
                    #mail attribute
                    send_email = self.send_email(
                                user= config.USERNAME_GMAIL,
                                pwd= config.PASSWORD_GMAIL,
                                recipient= self.R_EMAIL,
                              
                                subject= self.R_SUBJECT,
                                body= body_messages
                                )

                    mov, data = self.connection.uid('STORE', msg_uid, 
                                                    '+FLAGS', '(\Deleted)')

                    return False

                file = e.findexc(dir=config.BOOK_CHECK)
                os.remove(file)
                
                att_path = os.path.join(config.BOOK_CHECK, file_name)
                self.connection.uid('COPY', msg_uid, 'check_book')

                # finally write the stuff
                fp = open(att_path, 'wb')
                try:
                    fp.write(part_t.get_payload(decode=True))
                except:
                    continue

                fp.close()

                mov, data = self.connection.uid('STORE', msg_uid, 
                                                '+FLAGS', '(\Deleted)') 


                body_messages = (f"""Письмо с темой: {self.R_SUBJECT}.
Шаблон успешно обновлен:)
""")
                    #mail attribute
                send_email = self.send_email(
                                user= config.USERNAME_GMAIL,
                                pwd= config.PASSWORD_GMAIL,
                                recipient= self.R_EMAIL,
                                subject= self.R_SUBJECT,
                                body= body_messages
                                )

                return False

            else:
                
                att_path = os.path.join(config.ROOT_DIR, file_name)
                self.connection.uid('COPY', msg_uid, 'requests')

                # finally write the stuff
                fp = open(att_path, 'wb')
                try:
                    fp.write(part_t.get_payload(decode=True))
                except:
                    continue

                fp.close()

                mov, data = self.connection.uid('STORE', msg_uid, 
                                                '+FLAGS', '(\Deleted)') 

        return True

