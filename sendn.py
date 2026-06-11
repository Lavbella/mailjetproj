import base64
import csv
from mailjet_rest import Client

api_key = ''
api_secret = ''
mailjet = Client(auth=(api_key, api_secret), version='v3.1')

# 1. Ler os emails do ficheiro CSV
lista_emails = []
try:
    with open('contactos.csv', mode='r', encoding='utf-8-sig') as ficheiro:
        leitor = csv.reader(ficheiro)
        next(leitor)  # Salta a primeira linha ("email")
        for linha in leitor:
            if linha:  # Garante que a linha não está vazia
                email_limpo = linha[0].strip()
                lista_emails.append({"Email": email_limpo})
except FileNotFoundError:
    print("Erro: O ficheiro contactos.csv não foi encontrado.")
    exit()

# 2. Abrir e codificar o anexo (PDF)
with open("teste.pdf", "rb") as f:
    encoded_file = base64.b64encode(f.read()).decode('utf-8')

# 3. Montar a estrutura de envio com BCC
data = {
  'Messages': [
    {
      "From": {"Email": "noreply.xxxx@xxxx.pt", "Name": "xxxx"},
      "To": [{"Email": "noreply.xxxx@xxxx.pt"}], # Envia para ti mesmo no 'To'
      "Bcc": lista_emails, # Todos os do CSV ficam ocultos aqui
      "Subject": "Envio Geral com Anexo",
      "HTMLPart": "<h3>Olá, segue o ficheiro em anexo.</h3>",
      "Attachments": [
        {
          "ContentType": "application/pdf",
          "Filename": "teste.pdf",
          "Base64Content": encoded_file
        }
      ]
    }
  ]
}

# 4. Enviar
result = mailjet.send.create(data=data)
print(f"Status: {result.status_code}")
print(result.json())