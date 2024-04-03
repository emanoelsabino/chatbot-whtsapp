# imports necessarios no Módulo Utils.py
from Utils import *


nome_planilha = nome_planilha_inicial()  # declara o nome da planilha padrão
navegador = webdriver.Chrome()  # Configura o navegador e abre a aba do WhatsApp Web
navegador.get("https://web.whatsapp.com/")
verifica_conexao_whatsapp(navegador)  # Aguarda o WhatsApp Web abrir totalmente
# Loop para manter o WhatsApp Web aberto direto executando o chatbot
while True:
    try:
        op = '3'
        print('inicio...')
        hoje = datetime.now()  # Define a data de hoje
        variavel_erro = 'erro ao abrir planilha'  # variavel_erro -> aponta onde houve alguma exeção
        contatos = pd.read_excel(nome_planilha)
        tabela = load_workbook(nome_planilha)
        aba_ativa = tabela.active
        # Iteração na planilha para envio da mensagem
        for i, guia in enumerate(contatos["Número da guia"]):
            variavel_erro = f'erro ao criar variavel pessoa - posição {i+2}'
            pessoa = contatos.loc[i, "Nome"]
            variavel_erro = f'erro ao criar variavel telefone - pessoa {i+2} - {pessoa}'
            telefone = contatos.loc[i, "Telefone"]
            telefone_formatado = '55' + str(telefone)
            variavel_erro = f'erro ao criar variavel data_guia - pessoa {i+2} - {pessoa}'
            data_guia_str = contatos.loc[i, "Data de emissão"]
            mensagem = f'Sua guia nº {guia} vence em 5 dias, caso não for utilizar favor solicitar o cancelamento pelo link abaixo: \nhttps://docs.google.com/forms/d/e/1FAIpQLSfMFGuGujXeFpYrTu_fSrTuvfFs9e5QZq9NqwZrUoran0qlOw/viewform?vc=0&c=0&w=1&flr=0 \n\nCaso já tenha utilizado a guia, favor desconsiderar essa mensagem. \nA Formação Sanitária do 4º Batalhão de Engenharia de Combate agradece sua compreensão.'
            variavel_erro = f'erro ao converter data str para data - pessoa {i+2} - {pessoa}'
            data_guia = datetime.strptime(data_guia_str, '%d/%m/%Y')
            data_msg = data_guia + timedelta(days=25)
            print(data_msg)
            if data_msg.date() == hoje.date():
                if contatos.loc[i, "Msg Enviada"] != "ok":
                    if str(telefone) != 'nan':
                        # print('verificando para enviar msg')
                        print(pessoa, telefone_formatado)
                        tentativas = 0
                        while True:
                            try:
                                texto = urllib.parse.quote(f"Olá {pessoa}!\n{mensagem}\n\nEssa é uma menssagem automática.")
                                link = f"https://web.whatsapp.com/send?phone={telefone_formatado}&text={texto}"
                                time.sleep(2)
                                # print(link)
                                navegador.get(link)
                                time.sleep(3)
                                verifica_conexao_whatsapp(navegador)
                                time.sleep(30)
                                print('verifica numero valido ou invalido')
                                if len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) == 1:
                                    time.sleep(2)
                                    print('numero invalido')
                                    aba_ativa[f"A{i + 2}"] = "num_invalido"
                                    tabela.save(nome_planilha)
                                    time.sleep(2)
                                    break
                                
                                envia_msg_whatsapp(navegador)  # Envia mensagem / função do módulo Utils.py
                                time.sleep(7)
                                aba_ativa[f"A{i+2}"] = "ok"
                                tabela.save(nome_planilha)
                                print('msg enviada.')
                                break
                            except:
                                print('erro! Tentando novamente...')
                                tentativas += 1
                                if tentativas < 3:
                                    continue
                                else:
                                    break
        tabela.save(nome_planilha)
    except:
        print(f'{variavel_erro}')

    # Condição para nova execução ou finalizar o chatbot
    op, nome_planilha = condicao_novo_envio(hoje, nome_planilha, op)
    # Opção de finalizar o programa
    if op == 'sair':
        break
print('Programa encerrado!')
