import sys
import os
from mailjet_rest import Client
import streamlit as st
import pandas as pd
import base64
from docxtpl import DocxTemplate
#from docx2pdf import convert
import base64
import csv
import signal
import time
from streamlit.runtime import runtime
# IMPORTAÇÃO CORRETA: Mantém o módulo oficial sem avisos de erro
import streamlit.components.v1 as components
#import pythoncom


# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="xxxx Mail System", layout="wide")
st.title("📧 Gestão de Envios xxxx")

# 1. Inicializar a variável de controlo no estado da sessão
if "fechar_janela" not in st.session_state:
    st.session_state.fechar_janela = False


# Sidebar para Credenciais
with st.sidebar:
    st.header("Configurações API")
    # Carregar as chaves dos segredos (se não existirem, usa uma string vazia)
    api_key_default = st.secrets.get("MAILJET_API_KEY", "")
    api_secret_default = st.secrets.get("MAILJET_API_SECRET", "")

    # Os campos já aparecem preenchidos com os valores do ficheiro externo
    api_key = st.text_input("API Key", value=api_key_default, type="password")
    api_secret = st.text_input("API Secret", value=api_secret_default, type="password")
    st.divider()
    
    # Na barra lateral
    if st.sidebar.button("Sair e Encerrar"):
        st.session_state.fechar_janela = True
        

    # 3 OPÇÕES DE ENVIO
    opcao = st.radio(
        "Selecione o Modo de Envio:",
        ["1. CSV + PDF Único (BCC)", 
         "2. Mail Merge (Word + Excel)", 
         "3. Campanha via Template Mailjet",
         "4. CSV com corpo de email variavel"
         ]
    )

def gerar_html_corpo(saudacao_final_p, texto_principal):
    # Substitua pelo URL real do logótipo da xxxx se estiver online
    url_logo = "1757057210411.svg" 
    
    logo_base64 = carregar_logo_base64("EuroSorte.png")

    html = f"""
    <html>
    <body style="font-family: 'Open Sans'; color: #1a1a1a; line-height: 1.6; font-size: 10px;">
        <div style="max-width: 700px; margin: 0 auto; padding: 20px; border: 1px solid #f0f0f0;">
            
            <!-- Logótipo de Topo -->
            <div style="margin-bottom: 30px;">
                <img src="data:image/png;base64,{logo_base64}" alt="xxxx" style="max-height: 700px;">
            </div>

            <!-- Conteúdo -->
            <div style="margin-bottom: 40px;">
                <p style="font-family:'Open Sans', Arial, sans-serif; font-size:10px;" >{saudacao_final_p.replace('\n', '<br>')},</p>
                <br>
                <p style="font-family:'Open Sans', Arial, sans-serif; font-size:10px;" >{texto_principal.replace('\n', '<br>')}</p>
            </div>

            <!-- Assinatura -->
            <div style="border-top: 1px solid #eeeeee; padding-top: 15px; margin-top: 30px;">
                <p style="margin: 0;">Com os melhores cumprimentos,</p>
                <br>
                <p style="margin: 0;">O Presidente da Agência para a Gestão do Sistema xxxx</p>
                <p style="margin: 0;">xxxx xxxx xxxx</p>
            </div>

            <!-- Logótipo de Rodapé (Opcional ou versão reduzida) -->
            <div style="margin-top: 40px; opacity: 0.8;">
                <img src="{url_logo}" alt="xxxx" style="max-height: 40px;">
            </div>
            
        </div>
    </body>
    </html>
    """
    return html

def gerar_html_corpo_html(saudacao_final_p, texto_principal_html):
    # Substitua pelo URL real do logótipo da xxxx se estiver online
    url_logo = "1757057210411.svg" 
    
    logo_base64 = carregar_logo_base64("EuroSorte.png")

    html = f"""
    <html>
    <body style="font-family: 'Open Sans'; color: #1a1a1a; line-height: 1.6; font-size: 12px;">
        <div style="max-width: 700px; margin: 0 auto; padding: 20px; border: 1px solid #f0f0f0;">
            
            <!-- Logótipo de Topo -->
            <div style="margin-bottom: 30px;">
                <img src="data:image/png;base64,{logo_base64}" alt="xxxx" style="max-height: 700px;">
            </div>

            <!-- Conteúdo -->
            <div style="style="font-family:'Open Sans', Arial, sans-serif; font-size:12px; margin-bottom: 40px;">
                <p style="font-family:'Open Sans', Arial, sans-serif; font-size:12px;">{saudacao_final_p.replace('\n', '<br>')},</p>
                <br>
                {texto_principal_html}
            </div>

            <!-- Assinatura -->
            <div style="border-top: 1px solid #eeeeee; padding-top: 15px; margin-top: 30px;">
                <p style="margin: 0;">Com os melhores cumprimentos,</p>
                <br>
                <p style="margin: 0;">O Presidente da Agência para a Gestão do Sistema xxxx</p>
                <p style="margin: 0;">xxxx xxxx xxxx</p>
            </div>

            <!-- Logótipo de Rodapé (Opcional ou versão reduzida) -->
            <div style="margin-top: 40px; opacity: 0.8;">
                <img src="{url_logo}" alt="xxxx" style="max-height: 40px;">
            </div>
            
        </div>
    </body>
    </html>
    """
    return html

# Função auxiliar para codificar PDF
def get_base64_pdf(file_content):
    return base64.b64encode(file_content).decode('utf-8')

def carregar_logo_base64(caminho_imagem):
    with open(caminho_imagem, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode('utf-8')

def enviar_massa_bcc(api_key, api_secret, nome_csv, nome_pdf, assunto_p, corpo_p, saudacao_final_p):
    """
    Lê uma lista de emails de um CSV e envia um PDF único em BCC.
    """
    mailjet = Client(auth=(api_key, api_secret), version='v3.1')

    # 1. Ler os emails do ficheiro CSV
    lista_emails = []
    try:
        with open(nome_csv, mode='r', encoding='utf-8-sig') as ficheiro:
            leitor = csv.reader(ficheiro)
            next(leitor)  # Salta a primeira linha ("email")
            for linha in leitor:
                if linha:  # Garante que a linha não está vazia
                    email_limpo = linha[0].strip()
                    lista_emails.append({"Email": email_limpo})
    except FileNotFoundError:
            return {"erro": f"Ficheiro {nome_csv} não encontrado."}
    except Exception as e:
        return {"erro": f"Erro ao ler CSV: {str(e)}"}

     # 2. Abrir e codificar o anexo (PDF) fornecido
    try:
        with open(nome_pdf, "rb") as f:
            encoded_file = base64.b64encode(f.read()).decode('utf-8')
    except FileNotFoundError:
        return {"erro": f"Ficheiro {nome_pdf} não encontrado."}
    
    
    # Gerar o HTML final com a assinatura e logos
    html_final = gerar_html_corpo(saudacao_final_p, corpo_p)

    # 3. Montar a estrutura de envio com BCC
    data = {
    'Messages': [
        {
        "From": {"Email": "noreply.xxxx@xxxx.pt", "Name": "xxxx"},
        "To": [{"Email": "noreply.xxxx@xxxx.pt"}], # Envia para ti mesmo no 'To'
        "Bcc": lista_emails, # Todos os do CSV ficam ocultos aqui
        "Subject": assunto_p,
        "HTMLPart": html_final,
        "Attachments": [
            {
            "ContentType": "application/pdf",
            "Filename": nome_pdf,
            "Base64Content": encoded_file
            }
        ]
        }
    ]
    }

    # 4. Enviar
    result = mailjet.send.create(data=data)

    return {
        "status_code": result.status_code,
        "json": result.json()
    }

def enviar_massa_variaveis(
    api_key,
    api_secret,
    nome_csv,
    assunto,
    corpo_html_template,
    saudacao_final
):
    mailjet = Client(auth=(api_key, api_secret), version='v3.1')

    mensagens = []   # ← aqui vamos guardar TODOS os emails

    try:
        with open(nome_csv, mode='r', encoding='utf-8-sig') as ficheiro:
            leitor = csv.DictReader(ficheiro, delimiter=';')

            for linha in leitor:
                email_destino = linha.get("email", "").strip()
                if not email_destino:
                    continue

                # Substituir variáveis {coluna}
                try:
                    corpo_html = corpo_html_template.format(**linha)
                except KeyError as e:
                    return {
                        "status_code": 0,
                        "resposta": f"Variável inexistente no CSV: {e}"
                    }

                html_final = gerar_html_corpo_html(
                    saudacao_final,
                    corpo_html
                )

                mensagens.append({
                    "From": {
                        "Email": "noreply.xxxx@xxxx.pt",
                        "Name": "xxxx"
                    },
                    "To": [
                        {"Email": email_destino}
                    ],
                    "Subject": assunto,
                    "HTMLPart": html_final
                })

        if not mensagens:
            return {
                "status_code": -1,
                "resposta": "Nenhum email válido encontrado no CSV."
            }

        # ✅ Envio único com N mensagens
        data = {"Messages": mensagens}
        result = mailjet.send.create(data=data)

        if result.status_code >= 300:
            return {
                "status_code": -1,
                "resposta": result.json()
            }

    except FileNotFoundError:
        return {
                "status_code": -1,
                "resposta":f"Ficheiro {nome_csv} não encontrado."
            }

    except Exception as e:
        return {
                "status_code": -1,
                "resposta": str(e)
            }

    return {
        "status_code": 1,
        "resposta": str(len(mensagens))
    }


def get_main_dir():
    # Se estiver a correr como executável (.exe)
    if getattr(sys, 'frozen', False):
        # Retorna a pasta onde o run_app.exe reside
        return os.path.dirname(sys.executable)
    # Se estiver a correr como script (.py)
    return os.path.dirname(os.path.abspath(__file__))

def converter_docx_para_pdf(caminho_docx, pasta_destino):
    import subprocess
    import platform
    
    sistema = platform.system()
    # No Linux o comando é 'libreoffice', no Windows aponte para o .exe
    comando_bin = "libreoffice" if sistema != "Windows" else r"C:\Program Files\LibreOffice\program\soffice.exe"
    
    comando = [
        comando_bin,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", pasta_destino,
        caminho_docx
    ]
    
    # Executa a conversão
    subprocess.run(comando, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def Mail_Merge_Personalizado(api_key, api_secret, nome_xlsx, nome_docx, assunto_p, corpo_p, saudacao_final_p):
    # 0. Configurar caminhos
    RAIZ = get_main_dir()
    SUBPASTA = os.path.abspath(os.path.join(RAIZ, "output"))

    # Criar a subpasta se não existir para evitar erros de "caminho não encontrado"
    if not os.path.exists(SUBPASTA):
        os.makedirs(SUBPASTA)

    mailjet = Client(auth=(api_key, api_secret), version='v3.1')

    # 1. Ler Excel
    df = pd.read_excel(nome_xlsx)

    # Universal: substitui nulos por vazio para não aparecer 'nan' no Word/PDF
    df = df.fillna("")

    # Opcional: Garante que os nomes das colunas não têm espaços extras
    df.columns = [str(col).strip() for col in df.columns]

    for index, linha in df.iterrows():
        nome_base = f"documento_{index}"
        caminho_docx = os.path.abspath(os.path.join(SUBPASTA, f"{nome_base}.docx"))
        caminho_pdf = os.path.abspath(os.path.join(SUBPASTA, f"{nome_base}.pdf"))        
        
        # 2. Gerar Word
        doc = DocxTemplate(nome_docx)
        doc.render(linha.to_dict())
        doc.save(caminho_docx)
        
        # 3. Converter para PDF (Usando o LibreOffice conforme definido anteriormente)
        try:
            # IMPORTANTE: No LibreOffice, o segundo argumento deve ser a pasta de destino
            converter_docx_para_pdf(caminho_docx, SUBPASTA)
        except Exception as e:
            st.error(f"Erro na conversão do {nome_base}: {e}")
            continue

        # 4. Codificar para Base64 para o Mailjet
        # Adicionamos uma pequena verificação para garantir que o PDF existe antes de ler
        if os.path.exists(caminho_pdf):
            with open(caminho_pdf, "rb") as f:
                encoded_file = base64.b64encode(f.read()).decode('utf-8')
        else:
            st.error(f"Ficheiro PDF não gerado para o index {index}")
            continue
        
        html_final = gerar_html_corpo(saudacao_final_p, corpo_p)

        # 5. Enviar
        data = {
            'Messages': [
                {
                    "From": {"Email": "noreply.xxxx@xxxx.pt", "Name": "xxxx"},
                    "To": [{"Email": linha['Email']}],
                    "Subject": assunto_p,
                    "HTMLPart": html_final,
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

        if result.status_code != 200:
            return {"status": result.status_code, "json": result.json()}

    return "Processo concluído!"


def campanha_via_mailjet(api_key, api_secret, ID_contactos, ID_template, assunto_p, info_placeholder):

    mailjet = Client(auth=(api_key, api_secret), version='v3')

    # 1. OBTER O CONTEÚDO HTML DO TEMPLATE
    # O Mailjet guarda o HTML do template neste endpoint
    res_template = mailjet.template_detailcontent.get(id=ID_template)

    if res_template.status_code == 200:
        html_content = res_template.json()['Data'][0]['Html-part']
        info_placeholder.info("Conteúdo do template obtido com sucesso.")

        # 2. CRIAR O RASCUNHO (DRAFT)
        data_draft = {
            'Locale': "pt_PT",
            'Subject': assunto_p,
            'ContactsListID': ID_contactos,
            'SenderEmail': "noreply@xxxx.pt",
            'SenderName': "xxxx",
            'Title': "Envio_Final_xxxx"
        }
        res_draft = mailjet.campaigndraft.create(data=data_draft)

        if res_draft.status_code == 201:
            draft_id = res_draft.json()['Data'][0]['ID']
            info_placeholder.info(f"Rascunho criado: {draft_id}")

            # 3. INJETAR O HTML NO RASCUNHO
            # Aqui passamos o HTML extraído do template para o rascunho da campanha
            mailjet.campaigndraft_detailcontent.create(id=draft_id, data={'Html-part': html_content})

            # 4. ENVIAR A CAMPANHA
            res_send = mailjet.campaigndraft_send.create(id=draft_id)
            
            if res_send.status_code == 201:
                info_placeholder.info("Sucesso!...Campanha enviada para a lista com o template.")
                return 1
            else:
                info_placeholder.info(f"Erro no envio final: {res_send.json()}")
                return 0
        else:

            try:
                # Tenta transformar a resposta em texto legível
                detalhe_erro = res_draft.json() 
            except:
                # Se não for JSON, lê como texto simples
                detalhe_erro = res_draft.text

            info_placeholder.error(f"Erro no rascunho: {detalhe_erro}")
            return 0
    else:

        try:
            # Tenta transformar a resposta em texto legível
            detalhe_erro = res_template.json() 
        except:
            # Se não for JSON, lê como texto simples
            detalhe_erro = res_template.text

        info_placeholder.info(f"Erro ao obter template: {detalhe_erro}")
        return 0


# --- LÓGICA DAS OPÇÕES ---
if "1. CSV" in opcao:
    st.header("📤 Envio em Massa (BCC)")

    # 1. Procurar ficheiros na pasta do projeto
    ficheiros_csv = [f for f in os.listdir('.') if f.endswith('.csv')]
    ficheiros_pdf = [f for f in os.listdir('.') if f.endswith('.pdf')]

    # 2. Criar as caixas de seleção
    if not ficheiros_csv:
        st.warning("⚠️ Nenhum ficheiro CSV encontrado na pasta do projeto.")
    else:
        csv_file = st.selectbox("Selecione o ficheiro de contactos (CSV):", ficheiros_csv)

    if not ficheiros_pdf:
        st.warning("⚠️ Nenhum ficheiro PDF encontrado na pasta do projeto.")
    else:
        pdf_file = st.selectbox("Selecione o relatório (PDF):", ficheiros_pdf)
    
        # 1. Dropdown de Saudações
    opcoes_saudacao = [
        "Exmos. Senhores",
        "Exmo. Senhor Diretor",
        "Exma. Senhora Secretária-Geral",
        "Exmo. Senhor Presidente",
        "Personalizado..."
    ]
    
    escolha_saudacao = st.selectbox("Selecione a Saudação:", opcoes_saudacao)

    # Se escolher "Personalizado", abre um campo para escrever
    if escolha_saudacao == "Personalizado...":
        saudacao_final = st.text_input("Escreva a saudação personalizada:", value="Exmo. Senhor,")
    else:
        saudacao_final = escolha_saudacao

    assunto = st.text_input("Assunto", value="Documentação xxxx")
    corpo_input = st.text_area("Corpo do E-mail (HTML ou Texto simples)", 
                               value="Segue o ficheiro em anexo.", height=150)

    if st.button("Enviar para Todos"):
        # Lógica: Ler CSV -> Codificar PDF -> Mailjet Send v3.1 (Bcc)
        st.info("A processar envio em massa...")
        resultado = enviar_massa_bcc(api_key, api_secret, csv_file, pdf_file, saudacao_final, assunto, corpo_input)
        if resultado["status_code"] == 200:
            st.info("Envio em massa terminado com sucesso!...")
        else:
            st.info("Erro a enviar email em massa.") 

        st.info(resultado)

elif "2. Mail Merge" in opcao:
    st.header("📄 Mail Merge Personalizado")

    # 1. Procurar ficheiros na pasta do projeto
    ficheiros_xlsx = [f for f in os.listdir('.') if f.endswith('.xlsx')]
    ficheiros_docx = [f for f in os.listdir('.') if f.endswith('.docx')]

    # 2. Criar as caixas de seleção
    if not ficheiros_xlsx:
        st.warning("⚠️ Nenhum ficheiro xlsx encontrado na pasta do projeto.")
    else:
        xlsx_file = st.selectbox("Selecione o ficheiro de dados de configuração (xlsx):", ficheiros_xlsx)

    if not ficheiros_docx:
        st.warning("⚠️ Nenhum ficheiro docx encontrado na pasta do projeto.")
    else:
        docx_file = st.selectbox("Selecione o relatório (PDF):", ficheiros_docx)
    
        # 1. Dropdown de Saudações
    opcoes_saudacao = [
        "Exmos. Senhores",
        "Exmo. Senhor Diretor",
        "Exma. Senhora Secretária-Geral",
        "Exmo. Senhor Presidente",
        "Personalizado..."
    ]

    escolha_saudacao = st.selectbox("Selecione a Saudação:", opcoes_saudacao)

    # Se escolher "Personalizado", abre um campo para escrever
    if escolha_saudacao == "Personalizado...":
        saudacao_final = st.text_input("Escreva a saudação personalizada:", value="Exmo. Senhor,")
    else:
        saudacao_final = escolha_saudacao

    assunto = st.text_input("Assunto", value="Documentação xxxx")
    corpo_input = st.text_area("Corpo do E-mail (HTML ou Texto simples)", 
                               value="Segue o ficheiro em anexo.", height=150)


    if st.button("Gerar PDFs e Enviar"):
        # Lógica: Loop Excel -> DocxTemplate -> docx2pdf -> Mailjet Send
        st.warning("Nota: Requer Word instalado no servidor/PC local.")
        st.info("A processar...")
        resultado = Mail_Merge_Personalizado(api_key, api_secret, xlsx_file, docx_file, saudacao_final, assunto, corpo_input)
        st.info(resultado)
        

elif "3. Campanha" in opcao:

    st.header("🎨 Campanha via Mailjet")
    t_id = st.number_input("ID do Template Mailjet", step=1, value=111111)
    l_id = st.number_input("ID da Lista de Contactos", step=1, value=111111)
    subject_camp = st.text_input("Assunto da Campanha")

    if st.button("Enviar Campanha"):
        # Lógica: campaigndraft.create -> detailcontent -> send
        st.info("A processar a campanha...")
        resultado = campanha_via_mailjet(api_key, api_secret, l_id, t_id, subject_camp, st)
        if resultado > 0:
            st.success(f"Campanha {t_id} enviada de acordo com a lista {l_id}!")


if "4. CSV" in opcao:
    st.header("👉 Envio em Massa com corpo variável")

    # 1. Procurar ficheiros na pasta do projeto
    ficheiros_csv = [f for f in os.listdir('.') if f.endswith('.csv')]

    # 2. Criar as caixas de seleção
    if not ficheiros_csv:
        st.warning("⚠️ Nenhum ficheiro CSV encontrado na pasta do projeto.")
    else:
        csv_file = st.selectbox("Selecione o ficheiro de contactos (CSV):", ficheiros_csv)
    
        # 1. Dropdown de Saudações
    opcoes_saudacao = [
        "Exmos. Senhores",
        "Exmo. Senhor Diretor",
        "Exma. Senhora Secretária-Geral",
        "Exmo. Senhor Presidente",
        "Personalizado..."
    ]
    
    escolha_saudacao = st.selectbox("Selecione a Saudação:", opcoes_saudacao)

    # Se escolher "Personalizado", abre um campo para escrever
    if escolha_saudacao == "Personalizado...":
        saudacao_final = st.text_input("Escreva a saudação personalizada:", value="Exmo. Senhor,")
    else:
        saudacao_final = escolha_saudacao

    assunto = st.text_input("Assunto", value="Comunicação da xxxx")
    corpo_html_template = st.text_area("Corpo do E-mail (HTML ou Texto simples)", 
                               value="""
                            <p>Escola: {escola}</p>
                            <p>Por validar: {porvalidar}</p>

                            <p>
                            No âmbito xxxxxx 2026/2027, a xxxx informa que a aplicação da 2.ª xxxx encontra-se disponível
                            <a href="https://lavbella.com" target="_blank">aqui</a>.
                            </p>

                            <p><strong>Qual o prazo?</strong></p>

                            <p>
                            A aplicação encontra-se disponível até às 23h59 do dia 24 de abril, inclusive.
                            </p>

                            <p>
                            A xxxx permanece disponível para quaisquer esclarecimentos que considerem necessários e informa que foram elaboradas um conjunto
                            de FAQ para apoio ao xxxx xxxx, que podem ser consultadas
                            <a href="https://lavbella.com" target="_blank">aqui</a>.
                            </p>
                            """, height=150)

    if st.button("Enviar para Todos"):
        # Lógica: Ler CSV -> Codificar PDF -> Mailjet Send v3.1 (Bcc)
        st.info("A processar envio personalizado...")
        resultado = enviar_massa_variaveis(api_key, api_secret, csv_file, assunto, corpo_html_template, saudacao_final)
        if resultado["status_code"] >= 1:
            st.info("Envio personalizado de " + resultado["resposta"] + "emails terminado com sucesso!..." )
        else:
            st.info("Erro a enviar email personalizado:" + resultado["resposta"]) 

        st.info(resultado)

# --- RODAPÉ ---
st.divider()
st.caption("xxxx Mail System v1.0 - Desenvolvido em Python")

# ==========================================
# 3. FINAL DO FLUXO: Execução de Gatilhos
# ==========================================
# O bloco que gerencia o JavaScript fica solto aqui no final do script:
# 3. O bloco moderno que executa o comando (no fluxo principal, em baixo)
# 3. Execução estável no fluxo principal
if st.session_state.fechar_janela:
    # Cria uma sobreposição visual agressiva para "cegar" o ecrã laranja do Streamlit
    st.markdown("""
        <style>
            /* Bloqueia qualquer aviso de erro técnico do Streamlit de aparecer */
            div[data-testid="stConnectionStatus"], 
            .stStatusWidget, 
            #connection-fatal-error { display: none !important; }
            
            #modal-xxxx {
                position: fixed;
                top: 0; left: 0; width: 100vw; height: 100vh;
                background-color: white;
                z-index: 2147483647;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                font-family: sans-serif;
            }
        </style>
        <div id="modal-xxxx">
            <h1 style="color: #1E3A8A;">xxxx Mail Merge</h1>
            <p style="font-size: 1.2em;">Aplicação encerrada com sucesso.</p>
            <p style="color: #666;">A fechar processos... Pode fechar esta janela.</p>
        </div>
    """, unsafe_allow_html=True)

    # Executa o alerta e a limpeza de URL de forma segura no iframe oficial
    components.html(
        """
        <script>
            alert("A terminar processo xxxx...");
            window.parent.window.location.href = "about:blank";
            window.parent.window.close();
        </script>
        """,
        height=1, # 1 pixel para passar na validação estrita sem quebrar
    )
    
    # Pausa controlada para o utilizador aceitar o alerta antes do kill
    time.sleep(1.5)
    
    # Executa o kill para fechar o terminal Linux/Windows
    os.kill(os.getpid(), signal.SIGINT)
    
    st.session_state.fechar_janela = False