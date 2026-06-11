
import os
import base64
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
from mailjet_rest import Client

# Configurações
API_KEY = ''
API_SECRET = ''
SUBPASTA = "output"

# Criar a subpasta se não existir
if not os.path.exists(SUBPASTA):
    os.makedirs(SUBPASTA)

mailjet = Client(auth=(API_KEY, API_SECRET), version='v3.1')

# 1. Ler Excel
df = pd.read_excel('dados.xlsx')

for index, linha in df.iterrows():
    # Caminhos dos ficheiros dentro da subpasta
    nome_base = f"documento_{index}"
    caminho_docx = os.path.join(SUBPASTA, f"{nome_base}.docx")
    caminho_pdf = os.path.join(SUBPASTA, f"{nome_base}.pdf")
    
    # 2. Gerar Word a partir do Template
    doc = DocxTemplate("template.docx")
    doc.render(linha.to_dict())
    doc.save(caminho_docx)
    
    # 3. Converter para PDF (Guarda na subpasta)
    # Nota: No Windows, o docx2pdf precisa do caminho completo ou relativo correto
    convert(caminho_docx, caminho_pdf)
    
    # 4. Codificar para Base64 para o Mailjet
    with open(caminho_pdf, "rb") as f:
        encoded_file = base64.b64encode(f.read()).decode('utf-8')
    
    # 5. Enviar
    data = {
      'Messages': [
        {
          "From": {"Email": "noreply.xxxx@xxxx.xxx", "Name": "XXXX"},
          "To": [{"Email": linha['Email']}],
          "Subject": f"MARE",
          "HTMLPart": f"<p>Municipio de {linha['Concelho']}, o seu PDF está em anexo.</p>",
          "Attachments": [
            {
              "ContentType": "application/pdf",
              "Filename": f"{nome_base}.pdf",
              "Base64Content": encoded_file
            }
          ]
        }
      ]
    }
    
    result = mailjet.send.create(data=data)
    print(f"[{result.status_code}] Enviado para: {linha['Email']}")

    # 6. Limpeza (Opcional: remove os ficheiros da subpasta após envio)
    # os.remove(caminho_docx)
    # os.remove(caminho_pdf)

print("Processo concluído!")
