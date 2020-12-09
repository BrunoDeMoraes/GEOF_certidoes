from tkinter import *
from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time

tela = Tk()
tela.resizable(False, False)
tela.title('GEOF - Analisador de certidões')
#tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')

def executa():
    global variavel
    global variavel2
    global variavel3
    global obj1
    tempo_inicial = time.time()
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
    empresasdic = {}
    obj1 = Certidao(dia, mes, ano)
    obj1.mensagem_log('Início da execução')
    print("\nInício da Execução")
    lis_ref_cel = obj1.pega_referencia()
    if len(lis_ref_cel) == 0:
        obj1.mensagem_log('\nData específicada não encontrada')
        raise Exception('Data não encontrada!')
    elif len(lis_ref_cel) > 1:
        obj1.mensagem_log('Data informada em multiplicidade')
        print(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
                        f'\nApague os valores duplicados e execute o programa novamente.')
        raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
                        f'\nApague os valores duplicados e execute o programa novamente.')
    else:
        ref_cel = lis_ref_cel[0]
        obj1.mensagem_log(f'\nReferência encontrada na célula {lis_ref_cel[0]}')
    fornecedores = obj1.pega_fornecedores(ref_cel)
    fornecedores_cnpj = obj1.inclui_cnpj_em_fornecedores(fornecedores)
    lista_cnpj_completa = obj1.listar_cnpjs(fornecedores)
    obj1.mensagem_log_sem_horario('\nFornecedores analisados:')
    print('\nFornecedores analisados:')
    for emp in fornecedores:
        print(f'{emp}')
        obj1.mensagem_log_sem_horario(f'{emp}')
    obj1.cria_diretorio(fornecedores)
    obj1.apaga_imagem(fornecedores)
    obj1.certidoes_n_encontradas(fornecedores, orgaos)
    obj1.pdf_para_jpg(fornecedores, orgaos)
    objUniao = Uniao(dia, mes, ano, fornecedores)
    objTst = Tst(dia, mes, ano, fornecedores)
    objFgts = Fgts(dia, mes, ano, fornecedores)
    objGdf = Gdf(dia, mes, ano, fornecedores)
    lista_objetos = [objUniao, objTst, objFgts, objGdf]
    obj1.mensagem_log('\nInicio da conferência de datas de emissão e vencimento:')
    print('\nInicio da conferência de datas de emissão e vencimento:')
    print(f'\nTotal executado: {obj1.percentual}%')
    for emp in fornecedores:
        empresadic = {}
        index = 0
        print(f'\n{emp}')
        obj1.mensagem_log(f'\n{emp}')
        for objeto in lista_objetos:
            cert = objeto.pega_string(emp)
            obj1.percentual += (25 / len(fornecedores))
            print(f'\n    Total executado: {obj1.percentual}%')
            val, cnpj_para_comparação = objeto.confere_data(cert)
            if val == True and cnpj_para_comparação == fornecedores[emp][1]:
                empresadic[orgaos[index]] = 'OK'
            elif cnpj_para_comparação != fornecedores[emp][1]:
                empresadic[orgaos[index]] = 'CNPJ-ERRO'
            else:
                empresadic[orgaos[index]] = 'INCOMPATÍVEL'
            index += 1
        empresasdic[emp] = empresadic
    print('\nRESULTADO DA CONFERÊNCIA:')
    obj1.mensagem_log_sem_horario('\nRESULTADO DA CONFERÊNCIA:')
    numerador = 0
    for emp in empresasdic:
        print(f'{numerador + 1 :>2} - {emp}\n{empresasdic[emp]}\n')
        obj1.mensagem_log_sem_horario(f'{numerador + 1 :>2} - {emp}\n{empresasdic[emp]}\n')
        numerador += 1
    empresas_a_atualizar = {}
    for emp in empresasdic:
        certidoes_a_atualizar = []
        for orgao in empresasdic[emp]:
            if empresasdic[emp][orgao] == 'INCOMPATÍVEL' or empresasdic[emp][orgao] == 'CNPJ-ERRO':
                certidoes_a_atualizar.append(orgao)
        if len(certidoes_a_atualizar) > 0:
            empresas_a_atualizar[emp] = certidoes_a_atualizar
    obj1.pega_cnpj(empresas_a_atualizar)
    print('\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n')
    obj1.mensagem_log_sem_horario('\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n')
    for emp in empresas_a_atualizar:
        print(f'{emp} - {empresas_a_atualizar[emp][0:-1]} - CNPJ: {empresas_a_atualizar[emp][-1]}\n')
        obj1.mensagem_log_sem_horario(f'{emp} - {empresas_a_atualizar[emp][0:-1]} - CNPJ: {empresas_a_atualizar[emp][-1]}\n')
    obj1.apaga_imagem(fornecedores)
    tempo_final = time.time()
    tempo_de_execução = int((tempo_final - tempo_inicial))
    obj1.mensagem_log(
        f'\n\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
    obj1.mensagem_log_sem_horario(
        '\n\n======================================================================================\n')

def renomeia():
    global variavel
    global variavel2
    global variavel3
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    global obj1
    print('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES INICIADO:\n')
    obj1 = Certidao(dia, mes, ano)
    lis_ref_cel = obj1.pega_referencia()
    if len(lis_ref_cel) == 0:
        raise Exception('Data não encontrada!')
    elif len(lis_ref_cel) > 1:
        raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
                        f'\nApague os valores duplicados e execute o programa novamente.')
    else:
        ref_cel = lis_ref_cel[0]
    fornecedores = obj1.pega_fornecedores(ref_cel)
    obj1.apaga_imagem(fornecedores)
    obj1.pdf_para_jpg_renomear(fornecedores)
    obj1.gera_nome(fornecedores)
    obj1.apaga_imagem(fornecedores)

def transfere_certidoes():
    global variavel
    global variavel2
    global variavel3
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    global obj1
    obj1 = Certidao(dia, mes, ano)
    lis_ref_cel = obj1.pega_referencia()
    if len(lis_ref_cel) == 0:
        raise Exception('Data não encontrada!')
    elif len(lis_ref_cel) > 1:
        raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
                        f'\nApague os valores duplicados e execute o programa novamente.')
    else:
        ref_cel = lis_ref_cel[0]
    fornecedores = obj1.pega_fornecedores(ref_cel)
    obj1.certidoes_para_pagamento(fornecedores)

def mescla_certidoes():
    global variavel
    global variavel2
    global variavel3
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    global obj1
    obj1 = Certidao(dia, mes, ano)
    lis_ref_cel = obj1.pega_referencia()
    if len(lis_ref_cel) == 0:
        raise Exception('Data não encontrada!')
    elif len(lis_ref_cel) > 1:
        raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
                        f'\nApague os valores duplicados e execute o programa novamente.')
    else:
        ref_cel = lis_ref_cel[0]
    fornecedores = obj1.pega_fornecedores(ref_cel)
    obj1.merge(fornecedores)

frame_mestre = LabelFrame(tela, padx=0, pady=0)
frame_mestre.pack(padx=1, pady=1)

titulo = Label(frame_mestre, text ='''Indique a data limite pretendida para o próximo pagamento e
em seguida escolha uma das seguintes opções:''', pady=0, padx=0, bg='green', fg='white', font=('Helvetica', 12, 'bold'))

dia_etiqueta = Label(frame_mestre, text='Dia', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 9, 'bold'))
mes_etiqueta = Label(frame_mestre, text='Mês', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 9, 'bold'))
ano_etiqueta = Label(frame_mestre, text='Ano', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 9, 'bold'))

variavel = StringVar()
variavel.set(" ")
variavel2 = StringVar()
variavel2.set(" ")
variavel3 = StringVar()
variavel3.set(" ")
dias = [' ']
meses = [' ']
anos = [' ']
contador_dia = 1
while contador_dia <= 31:
    if contador_dia < 10:
        dias.append(f"0{contador_dia}")
        contador_dia += 1
    else:
        dias.append(str(contador_dia))
        contador_dia += 1
contador_mes = 1
while contador_mes <= 12:
    if contador_mes < 10:
        meses.append(f"0{contador_mes}")
        contador_mes += 1
    else:
        meses.append(str(contador_mes))
        contador_mes += 1
contador_anos = 2010
while contador_anos <= 2040:
    anos.append(str(contador_anos))
    contador_anos += 1

validacao = OptionMenu(frame_mestre, variavel, *dias)
validacao2 = OptionMenu(frame_mestre, variavel2, *meses)
validacao3 = OptionMenu(frame_mestre, variavel3, *anos)

titulo_analisar = Label(frame_mestre, text ='''Utilize esta opção para identificar quais certidões devem ser atualizadas
ou se há requisitos a cumprir para a devida execução da análise.''', pady=0, padx=0, bg='white',
                        fg='black', font=('Helvetica', 9, 'bold'))

botao_analisar = Button(frame_mestre, text='Analisar\ncertidões', command=executa, padx=25, pady=10, bg='green',
                        fg='white', font=('Helvetica', 11, 'bold'), bd=1)

titulo_renomear = Label(frame_mestre, text ='''Após atualizar as certidões, use esta opção para padronizar os nomes dos 
arquivos e em seguida faça nova análise para certificar que está tudo OK.''', pady=0, padx=0, bg='white', fg='black', font=('Helvetica', 9, 'bold'))

botao_renomear = Button(frame_mestre, text='Renomear\ncertidões', command=renomeia, padx=20, pady=10, bg='green',
                        fg='white', font=('Helvetica', 11, 'bold'), bd=1)

titulo_transfere_arquivos = Label(frame_mestre, text ='''Esta opção transfere as certidões para uma pasta identificada pela data
do pagamento. Esse passo deve ser executado logo após a análise.''', pady=0, padx=0, bg='white', fg='black', font=('Helvetica', 9, 'bold'))

botao_transfere_arquivos = Button(frame_mestre, text='Transferir\ncertidões', command=transfere_certidoes, padx=20, pady=10, bg='green',
                        fg='white', font=('Helvetica', 11, 'bold'), bd=1)

titulo_mescla_arquivos = Label(frame_mestre, text ='''Se já houve o pagamento e os comprovantes estão na devida pasta, esta 
opção mescla os comprovantes com suas respectivas certidões.''', pady=0, padx=0, bg='white', fg='black', font=('Helvetica', 9, 'bold'))

botao_mescla_arquivos = Button(frame_mestre, text=' Mesclar  \narquivos', command=mescla_certidoes, padx=20, pady=10, bg='green',
                        fg='white', font=('Helvetica', 11, 'bold'), bd=1)

roda_pe = Label(frame_mestre, text ="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
               font=('Helvetica', 8, 'italic'), anchor=E)

titulo.grid(row=0, column=1, columnspan=5, pady=10, sticky=W+E)
dia_etiqueta.grid(row=1, column=1, pady=0, ipadx=14, ipady=0)
mes_etiqueta.grid(row=1, column=2, pady=0, ipadx=14, ipady=0)
ano_etiqueta.grid(row=1, column=3, pady=0, ipadx=14, ipady=0)
validacao.grid(row=2, column=1, pady=0)
validacao2.grid(row=2, column=2, pady=0)
validacao3.grid(row=2, column=3, pady=0)
titulo_analisar.grid(row=3, column=4, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
botao_analisar.grid(row=3, column=5, padx=0, pady=10)
titulo_renomear.grid(row=4, column=4, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
botao_renomear.grid(row=4, column=5, padx=0,pady=10)
titulo_transfere_arquivos.grid(row=5, column=4, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
botao_transfere_arquivos.grid(row=5, column=5, padx=0,pady=10)
titulo_mescla_arquivos.grid(row=6, column=4, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
botao_mescla_arquivos.grid(row=6, column=5, padx=0,pady=10)
roda_pe.grid(row=7, column=1, columnspan=10, pady=5, sticky=W+E)

tela.mainloop()
