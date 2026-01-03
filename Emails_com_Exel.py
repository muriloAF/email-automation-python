import openpyxl
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#Email parts

#server_smtp = "smtp.gmail.com"
#port = 587
#sender_email = "emailenvio@gmail.com.br" Exemplo se for usar Gmail
#password = "jvqesdxndauwtlyrdh" A senha se transforma no token gerado pelo google para apps menos seguros


#Configurações do servidor 
server_smtp = "smtp.zoho.com" #Servidor SMTP do Zoho Mail, muda de acordo com o serviço de envio de email utilizados
port = 587 #Porta do servidor SMTP, também muda de acordo com o serviço de envio de emails
sender_email = "emailenvio@gmail.com.br" #email de envio
password = "senha" #senha do email


#EXEL PARTS
#Loading the file
book = openpyxl.load_workbook('Clientes_email.xlsx')
#Selectin the page
emails_page = book['Plan1']
#printing each line data

subject = "Proposta ACM Grupos Geradores"
for i in emails_page.iter_rows(min_row=52, max_row=388):
    reciver_names = i[0].value
    reciver_email = i[5].value
    body = """
    <!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Proposta ACM Grupos Geradores</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f1f5f9;
            color: #333;
            line-height: 1.6;
        }}
        .container {{
            max-width: 600px;
            margin: 30px auto;
            background: #fff;
            padding: 25px;
            border: 1px solid #ccc;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }}
        .header {{
            text-align: center;
            margin-bottom: 25px;
        }}
        .header img {{
            max-width: 100%;
            border-radius: 8px;
            margin-bottom: 15px;
        }}
        .header h1 {{
            font-size: 22px;
            color: #004aad;
            margin: 0;
            text-transform: uppercase;
        }}
        .content h2 {{
            font-size: 20px;
            color: #004aad;
            margin-bottom: 15px;
            border-bottom: 2px solid #004aad;
            display: inline-block;
            padding-bottom: 5px;
        }}
        .content p {{
            margin-bottom: 12px;
            text-align: justify;
        }}
        ul {{
            margin: 15px 0;
            padding-left: 20px;
        }}
        ul li {{
            margin-bottom: 8px;
        }}
        .signature {{
            margin-top: 30px;
            border-top: 2px solid #004aad;
            padding-top: 20px;
            font-size: 14px;
            color: #555;
        }}
        .signature a {{
            color: #004aad;
            text-decoration: none;
            font-weight: bold;
        }}
        .footer {{
            margin-top: 30px;
            font-size: 12px;
            text-align: center;
            color: #777;
        }}
        .footer a {{
            color: #004aad;
            text-decoration: underline;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>ACM Grupos Geradores</h1>
        </div>
        <div class="content">
            <h2>Proposta Comercial</h2>
            <p>À equipe responsável pelos geradores da {name},</p>
            <p>É com grande satisfação que apresentamos a ACM Grupos Geradores, uma empresa especializada na manutenção e fornecimento de soluções de alta performance para grupos geradores. Nosso compromisso é garantir que sua operação nunca fique sem energia.</p>
            <p>Oferecemos:</p>
            <ul>
                <li>Manutenção preventiva e corretiva de grupos geradores;</li>
                <li>Higienização de tanques de diesel;</li>
                <li>Profissionais altamente qualificados e equipamentos de última geração.</li>
            </ul>
            <p>Caso tenha interesse, estamos à disposição para agendar uma reunião ou fornecer mais detalhes sobre nossa proposta.</p>
            <p>Agradecemos pela oportunidade de apresentar nossos serviços e esperamos colaborar com sua empresa.</p>
        </div>
        <div class="signature">
            <strong>Adriano Luiz da Fonseca</strong><br>
            Tecnólogo Mecatrônico<br>
            <a href="mailto:contato@acmgruposgeradores.com.br">contato@acmgruposgeradores.com.br</a><br>
            <a href="mailto:admfeitosa@acmgruposgeradores.com.br">admfeitosa@acmgruposgeradores.com.br</a>
        </div>
        <div class="footer">
            &copy; 2025 ACM Grupos Geradores. Todos os direitos reservados.<br>
            <a href="https://www.acmgruposgeradores.com.br">Visite nosso site</a>
        </div>
    </div>
</body>
</html>

    """.format(name=reciver_names)
    time.sleep(20)
    print(f'envio {reciver_names} para o email {reciver_email}')
    message = MIMEMultipart()
    message ['From'] = sender_email
    message ['To'] = reciver_email
    message ['Subject'] = subject
    message.attach(MIMEText(body, "html"))

    #Conecting SMTP server
    try:
        server = smtplib.SMTP(server_smtp, port)
        server.starttls()
        server.login(sender_email, password) 
        server.sendmail(sender_email, reciver_email, message.as_string())
        print("Sended")
    except Exception as e:
        print(f'Error: {e}')
    finally:
        server.quit()



                                                                       
