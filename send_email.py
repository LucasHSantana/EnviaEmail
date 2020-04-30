import os
import smtplib
import imghdr
from email.message import EmailMessage
import ssl

smtp_list= {
    'gmail': ['ssl', 'smtp.gmail.com'],
    'outlook': ['tls', 'smtp.live.com'],
    'office365': ['tls', 'smtp.office365.com'],
    'yahoomail': ['ssl', 'smtp.mail.yahoo.com'],
    'yahoomailplus': ['ssl', 'plus.smtp.mail.yahoo.com'],
    'yahoouk': ['ssl', 'smtp.mail.yahoo.co.uk'],
    'yahoodeutschland': ['ssl', 'smtp.mail.yahoo.com'],
    'yahooaunz': ['ssl', 'smtp.mail.yahoo.com.au'],
    'o2': ['none', 'smtp.o2.ie'],
    'o2uk': ['none', 'smtp.o2.co.uk'],
    'aol': ['tls', 'smtp.aol.com'],
    'at&t': ['ssl', 'smtp.att.yahoo.com'],
    'ntl': ['ssl', 'smtp.ntlworld.com'],
    'btconnect': ['none', 'pop3.btconnect.com'],
    'btopenworld': ['none', 'mail.btopenworld.com'],
    'btinternet': ['none', 'mail.btinternet.com'],
    'orange': ['none', 'smtp.orange.net'],
    'orangeuk': ['none', 'smtp.orange.co.uk'],
    'wanadoo': ['none', 'smtp.wanadoo.co.uk'],
    'hotmail': ['ssl', 'smtp.live.com'],
    'o2onlinedeutschland': ['none', 'mail.o2online.de'],
    'verizon': ['ssl', 'outgoing.verizon.net'],
    'verizon_yahoo': ['tls', 'outgoing.yahoo.verizon.net'],
    'mail.com': ['tls', 'smtp.mail.com'],
    'teste': ['none', 'localhost'],
}

SSL_PORT = 465
TLS_PORT = 587
DEFAULT_PORT = 25

class Email:    
    __emails_to = []
    __subject = ''
    __content = ''    
    __file_data = []
    __file_type = []
    __file_name = []

    def __init__(self, email='', passwd='', smtp=''):
        self.__login = email
        self.__passwd = passwd    
        self.__smtp = smtp

    '''
    Caso email e senha de login não sejam informados ao instanciar a classe
    pode-se informar através do set_from
    '''    
    def login(self, email, passwd, stmp):
        self.__login = email
        self.__passwd = passwd   
        self.__smtp = stmp

    #Retorna uma lista com os emails de destino informados
    @property
    def emails_to(self):        
        if type(self.__emails_to) is str:            
            l = self.__emails_to.split(',')
            l = [i.strip() for i in l]
            return l
        else:
            return self.__emails_to 

    '''
    Informar email(s) de destino, pode-se informar de três formas:
    string com email único: 'teste@teste.com'
    string com multiplos emails separados por vírgula: 'teste@teste.com, teste2@teste.com'
    lista: ['teste@teste.com', 'teste2@teste.com']
    '''
    @emails_to.setter
    def emails_to(self, emails):
        self.__emails_to = emails

    #Retorna o assunto informado
    @property
    def subject(self):
        return self.__subject

    #Informar Assunto do email(string)
    @subject.setter
    def subject(self, subject):
        self.__subject = subject

    #Retorna o conteúdo informado
    @property
    def content(self):
        return self.__content

    #Informar conteúdo do email(string ou texto html)
    @content.setter
    def content(self, content):
        self.__content = content

    #Retorna o email de login informado
    def get_login(self):
        return self.email_from

    '''
    Informar anexo, precisa ser uma lista de strings do caminho onde os anexos se encontram:
    ex: ['C:\imagem\imagem1.jpg', 'C:\imagem\imagem2.jpg']
    '''

    def attachment(self, filenames):
        supported_types = ['jpg', 'jpeg', 'png', 'gif', 'pdf', 'txt', 'py', 'doc', 'xls']        
        
        for filename in filenames:
            with open(filename, 'rb') as f:
                file_type = f.name.split('.')[-1] #Pega a extensão do arquivo

                if not file_type in supported_types:                
                    raise Exception('Tipo de arquivo não suportado!')

                self.__file_data.append(f.read())
                self.__file_type.append(file_type)
                self.__file_name.append(f.name.split('\\')[-1])
                f.close()                                   

    #Envia o email
    def send_email(self):                           
        if len(self.__login) < 6 or len(self.__passwd) <= 0:
            raise Exception('Informe o login e senha corretamente!')

        if len(self.__emails_to) <= 0:
            raise Exception('Informe um destinatário!')

        if len(self.__subject) <= 0:
            raise Exception('Informe um assunto!')

        msg = EmailMessage()
        msg['From'] = self.__login
        msg['To'] = self.__emails_to
        msg['Subject'] = self.__subject  

        #Verifica se o conteúdo do email é um html        
        if self.__content[:5] == '<html' or self.__content[:9] == '<!DOCTYPE':
            msg.add_alternative(self.__content, subtype='html')
        else:
            msg.set_content(self.__content)

        for file, file_type, file_name in zip(self.__file_data, self.__file_type, self.__file_name):
            if file_type in ['jpg', 'jpeg', 'png', 'gif']:
                maintype = 'image'
            else:
                maintype = 'application'

            msg.add_attachment(file, maintype=maintype, subtype=file_type, filename=file_name)

        if len(self.__smtp) <= 0:
            raise Exception('Informe um stmp a ser utilizado')

        if not self.__smtp in smtp_list:
            raise Exception('O smtp informado não é suportado!')

        context = ssl.create_default_context()

        if smtp_list[self.__smtp][0] == 'ssl':
            smtp = smtplib.SMTP_SSL(smtp_list[self.__smtp][1], SSL_PORT, context=context)
            
        elif smtp_list[self.__smtp][0] == 'tls':
            smtp = smtplib.SMTP(smtp_list[self.__smtp][1], TLS_PORT)
            smtp.starttls()                
        else:
            smtp = smtplib.SMTP(smtp_list[self.__smtp][1], DEFAULT_PORT)                    

        smtp.login(self.__login, self.__passwd)
        smtp.send_message(msg)

        smtp.quit()
    
    #Visualização
    def __str__(self):
        return f'From: {self.__login}\nTo: {self.__emails_to}\nSubject: {self.__subject}\nContent: {self.__content}'

if __name__ == '__main__':
    email = Email('lucas.hesantana16@gmail.com', '5Fk4@tmP@^Xf', 'gmail')
    email.emails_to = 'lucas.hesantana16@gmail.com'
    email.subject = 'Teste'    
    email.content = '<!DOCTYPE html><html><body><ul><li>Teste 1</li><li>Teste 2</li></ul></body></html>'    
    # email.attachment(['E:\Documentos\Livros\Livros Gratuitos em inglês\An Anthology of London in Literature, 1558-1914.pdf'])
    # print(email)
    email.send_email()
    