# Envio automatico pelo Gmail
# Configurações de e-mail
de = 'leonardorfragoso@gmail.com'
para = 'leonardorfragoso2@gmail.com'
senha = 'ppwbgavdaaaowxeo'
assunto = 'Bases compiladas'
mensagem = 'Automação para Compilar bases de dados em uma unica planila de Excel e enviar automaticamente por email.'

# Configuração do servidor SMTP do Gmail
servidor_smtp = 'smtp.gmail.com'
porta = 587

# Crie uma mensagem de e-mail
mensagem_email = MIMEMultipart()
mensagem_email['From'] = de
mensagem_email['To'] = para
mensagem_email['Subject'] = assunto
mensagem_email.attach(MIMEText(mensagem, 'plain'))

# Anexando um arquivos
caminho = os.getcwd()
arquivo_anexo = os.path.join(caminho, "Vendas.xlsx")
with open(arquivo_anexo, 'rb') as arquivo:
    part = MIMEApplication(arquivo.read(), _subtype="pdf")
part.add_header('Content-Disposition', 'attachment', filename=arquivo_anexo)
mensagem_email.attach(part)

# Conectando ao servidor SMTP e enviando o e-mail
try:
    servidor = smtplib.SMTP(servidor_smtp, porta)
    servidor.starttls()
    servidor.login(de, senha)
    servidor.sendmail(de, para, mensagem_email.as_string())
    servidor.quit()
    print('E-mail enviado com sucesso')
except Exception as e:
    print('Erro ao enviar o e-mail:', str(e))
