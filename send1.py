import base64
from mailjet_rest import Client

api_key = 'xxxx'
api_secret = 'xxxx'
mailjet = Client(auth=(api_key, api_secret), version='v3.1')

# Abrir e codificar o ficheiro
with open("teste.pdf", "rb") as f:
    encoded_file = base64.b64encode(f.read()).decode('utf-8')

data = {
  'Messages': [
    {
      "From": {"Email": "noreply.xxx@xxx.pt", "Name": "xxxx"},
      "To": [{"Email": "xxx@gmail.com"}],
      "Subject": "Envio com Anexo",
      "HTMLPart": "<h3>Segue o ficheiro em anexo.</h3>",
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

result = mailjet.send.create(data=data)
print(result.status_code)
print(result.json()) # Isto vai dizer exatamente por que falhou