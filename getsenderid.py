from mailjet_rest import Client
mailjet = Client(auth=('', ''), version='v3')
result = mailjet.sender.get()
print(result.json()) # Procura o campo "ID" no texto que aparecer