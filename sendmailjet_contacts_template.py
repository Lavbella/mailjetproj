from mailjet_rest import Client

api_key = 'xxxx'
api_secret = 'xxxx'
mailjet = Client(auth=(api_key, api_secret), version='v3')

TEMPLATE_ID = 1111
LIST_ID = 1111

# 1. OBTER O CONTEÚDO HTML DO TEMPLATE
# O Mailjet guarda o HTML do template neste endpoint
res_template = mailjet.template_detailcontent.get(id=TEMPLATE_ID)

if res_template.status_code == 200:
    html_content = res_template.json()['Data'][0]['Html-part']
    print("Conteúdo do template obtido com sucesso.")

    # 2. CRIAR O RASCUNHO (DRAFT)
    data_draft = {
        'Locale': "pt_PT",
        'Subject': "Assunto da Campanha",
        'ContactsListID': LIST_ID,
        'SenderEmail': "noreply@xxxx.pt",
        'SenderName': "xxxx",
        'Title': "Envio_Final_xxxx"
    }
    res_draft = mailjet.campaigndraft.create(data=data_draft)

    if res_draft.status_code == 201:
        draft_id = res_draft.json()['Data'][0]['ID']
        print(f"Rascunho criado: {draft_id}")

        # 3. INJETAR O HTML NO RASCUNHO
        # Aqui passamos o HTML extraído do template para o rascunho da campanha
        mailjet.campaigndraft_detailcontent.create(id=draft_id, data={'Html-part': html_content})

        # 4. ENVIAR A CAMPANHA
        res_send = mailjet.campaigndraft_send.create(id=draft_id)
        
        if res_send.status_code == 201:
            print("Sucesso! Campanha enviada para a lista com o template.")
        else:
            print(f"Erro no envio final: {res_send.json()}")
    else:
        print(f"Erro no rascunho: {res_draft.json()}")
else:
    print(f"Erro ao obter template: {res_template.json()}")