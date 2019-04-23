# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import csv


#define o driver para ser utilizado com o Chrome
driver = webdriver.Chrome()

#função para coletar os dados nas páginas das requisições, após a página com os dados de cada requisição ser aberta
def coleta_dados():
    #pega todas as linhas da tabela da página da requisição
    linhas_tabela = driver.find_elements(By.TAG_NAME, 'td')
    #cria uma lista vazia de dados da tabela
    lista_dados_tabela = []
    #faz um loop para preencher a lista de dados da tabela com as informações de casa célula da tabela
    for linha in linhas_tabela:
        #insere os dados de cada célula da tabela na lista criada
        lista_dados_tabela.append(linha.text)
    #cria as variáveis que serão utilizadas para escrever os dados no arquivo csv
    tipo_requisicao_completo = lista_dados_tabela[0]
    global tipo_requisicao
    global numero_requisicao
    global solicitante
    global data_requisicao
    global percurso
    global tipo_veiculo
    global custo_estimado
    global custo_real
    global finalidade

    #como há variações nas tabelas de dados de acordo com o tipo de requisição:
    #seleciona as requisições do tipo RQT1
    if tipo_requisicao_completo == 'Requisição: Tipo RQT1':
        tipo_requisicao = tipo_requisicao_completo[17:]
        numero_requisicao_geral = lista_dados_tabela[1]
        numero_requisicao = numero_requisicao_geral[22:28]
        solicitante = lista_dados_tabela[16]
        data_requisicao = lista_dados_tabela[25]
        percurso = lista_dados_tabela[47]
        tipo_veiculo = lista_dados_tabela[41]
        custo_estimado = lista_dados_tabela[43]
        custo_real = lista_dados_tabela[45]
        finalidade = lista_dados_tabela[49]

    # seleciona as requisições do tipo RQT2
    elif tipo_requisicao_completo == 'Requisição: Tipo RQT2':
        tipo_requisicao = tipo_requisicao_completo[17:]
        numero_requisicao_geral = lista_dados_tabela[1]
        numero_requisicao = numero_requisicao_geral[22:28]
        solicitante = lista_dados_tabela[16]
        data_requisicao = lista_dados_tabela[25]
        percurso = lista_dados_tabela[51]
        tipo_veiculo = lista_dados_tabela[45]
        custo_estimado = lista_dados_tabela[47]
        custo_real = lista_dados_tabela[49]
        finalidade = lista_dados_tabela[53]

    # seleciona as requisições do tipo terceirizado
    elif tipo_requisicao_completo == 'Requisição: Tipo Terceirizado':
        tipo_requisicao = tipo_requisicao_completo[17:]
        numero_requisicao_geral = lista_dados_tabela[1]
        numero_requisicao = numero_requisicao_geral[22:28]
        solicitante = lista_dados_tabela[12]
        data_requisicao = lista_dados_tabela[21]
        percurso = lista_dados_tabela[47]
        tipo_veiculo = lista_dados_tabela[41]
        custo_estimado = lista_dados_tabela[43]
        custo_real = lista_dados_tabela[45]
        finalidade = lista_dados_tabela[49]

#cria o arquivo transporte.csv, no modo escrita w (writer)
with open('transporte.csv', 'w') as arquivo:
    #cria a variável writer para escrever no arquivo
    writer = csv.writer(arquivo)
    #escreve as colunas que serão utilizadas como cabeçalho do arquivo csv
    writer.writerow(['Número da requisição', 'Tipo de requisição', 'Data', 'Percurso', 'Custo estimado', 'Custo real', 'Veículo', 'Solicitante', 'Finalidade'])


#função para acessar o siscore
def acesso_siscore():
    #acessa a página principal do siscore
    driver.get("https://www2.dti.ufv.br/dtr_siscore/")
    #procura os elementos e envia as credenciais de acesso
    driver.find_element_by_id("Usuario").click()
    driver.find_element_by_id("Usuario").clear()
    usuario = str(input('Informe a matrícula: '))
    driver.find_element_by_id("Usuario").send_keys(usuario)
    driver.find_element_by_id("Senha").click()
    driver.find_element_by_id("Senha").clear()
    senha = str(input('Informe a senha: '))
    driver.find_element_by_id("Senha").send_keys(senha)
    #faz o login no site do siscore
    driver.find_element_by_id("Login").click()
    #aguarda alguns segundos para o site carregar após o login
    sleep(5)
    #procura o elemento Requisição para acessar a página de requisições
    driver.find_element_by_link_text(u"Requisição").click()
    #busca as requisições do ano atual
    driver.find_element_by_link_text("Consultar/Atualizar").click()
    #clica no botão Consulta para abrir as requisições cadastradas no ano atual
    driver.find_element_by_name("btnSubmit").click()
    #aguarda 10 segundos para que todos os dados da página de requisições sejam carregados
    sleep(10)
    #procura todos os ícones da página de requisições, para incluí-los em uma lista
    icones_pagina = driver.find_elements(By.XPATH, "//body/div/div/table/tbody/tr/td/a/img")
    #cria a lista vazia de requisições
    lista_requisicoes_driver = []
    #faz um loop nos ícones da página para obter somente o ícone que tem o link de acesso à página com os dados de casa requisição
    for elemento in icones_pagina:
        #se o atributo src do elemento (webdriver) for correspondente ao ícone que tem o link de acesso à página de cada requisição
        if elemento.get_attribute('src') == 'https://www2.dti.ufv.br/dtr_siscore/images/btn_search_bg.gif':
            #inclui o elemento (webdriver) na lista de requisições a serem visitadas para coleta de dados (scrap)
            lista_requisicoes_driver.append(elemento)

    #começa na 3ª linha da tabela de dados do siscore, que corresponde à lista de requisições já cadastradas no sistema, linhas 1 e 2 são apenas cabeçalhos da tabela
    linha = 3
    #faz um loop para acessar cada requisição na lista de requisições criadas anteriormente
    for requisicao in lista_requisicoes_driver:
            # define a página atual com handles[0]
            window_berofe = driver.window_handles[0]
            #procura o ícone de acesso à página de cada requisição, que contêm os dados a serem raspados (scrap)
            driver.find_element_by_xpath(f'//*[@id="estiloTabela"]/tbody/tr[{linha}]/td[10]/a/img').click()
            # define que o driver vá para a nova página aberta, definida como handles[1]
            window_after = driver.window_handles[1]
            #muda o foco do driver para a página da requisição que foi aberta
            driver.switch_to_window(window_after)
            #executa a função para coletar os dados na página da requisição
            coleta_dados()
            #atualiza o arquivo transporte csv com os dados coletados/raspados (scrap)
            with open('transporte.csv', 'a') as arquivo:
                writer = csv.writer(arquivo)
                # função para escrever os dados coletados em um arquivo do tipo csv
                writer.writerow([numero_requisicao, tipo_requisicao, data_requisicao, percurso, custo_estimado, custo_real, tipo_veiculo, solicitante, finalidade])

            #fecha o drive e a página de requisição, voltando para a página que contêm todas as requisições
            driver.close()
            #muda o foco do driver para a página que contêm todas as requisições
            driver.switch_to_window(window_berofe)
            #atualiza a linha da tabela que contêm todas as requisições, para abrir a página da próxima requisição, que ainda não foi aberta
            linha +=1

#executa a função de acesso ao siscore
acesso_siscore()
#fecha o arquivo csv com os dados gravados
arquivo.close()
#fecha o driver para concluir o processo
driver.close()