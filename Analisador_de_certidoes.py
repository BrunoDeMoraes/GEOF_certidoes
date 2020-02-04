from tkinter import *
from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time
from PIL import ImageTk, Image

tela = Tk()
tela.resizable(False, False)
tela.title('GEOF - Analisador de certidões')
tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')

def executa():
    global variavel
    global variavel2
    global variavel3
    tempo_inicial = time.time()
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
    empresasdic = {}

    global obj1
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
    obj1.mensagem_log_sem_horario('\nFornecedores analisados:')
    print('\nFornecedores analisados:')
    for emp in fornecedores:
        print(f'{emp}')
        obj1.mensagem_log_sem_horario(f'{emp}')
    obj1.cria_diretorio(fornecedores)
    obj1.certidoes_n_encontradas(fornecedores, orgaos)
    obj1.pdf_para_jpg(fornecedores, orgaos)
    objUniao = Uniao(dia, mes, ano)
    objTst = Tst(dia, mes, ano)
    objFgts = Fgts(dia, mes, ano)
    objGdf = Gdf(dia, mes, ano)
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
            val = objeto.confere_data(cert)
            if val == True:
                empresadic[orgaos[index]] = 'OK'
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
    obj1.apaga_imagem(fornecedores)
    tempo_final = time.time()
    tempo_de_execução = int((tempo_final - tempo_inicial))
    obj1.mensagem_log(
        f'\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
    obj1.mensagem_log_sem_horario(
        '\n\n======================================================================================\n')

def renomeia():
    global variavel
    global variavel2
    global variavel3
    tempo_inicial = time.time()
    dia = variavel.get()
    mes = variavel2.get()
    ano = variavel3.get()
    orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
    empresasdic = {}

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
    print(fornecedores)
    obj1.pdf_para_jpg_renomear(fornecedores)
    obj1.gera_nome(fornecedores)
    obj1.apaga_imagem(fornecedores)

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

frame_2 = LabelFrame(tela, padx=0, pady=0)
frame_2.pack(padx=1, pady=1)

titulo = Label(frame_2,  text = "Escolha a data pretendida para o próximo pagamento\ne em seguida pressione"
                                       " o botão 'Analisar certidões'.", pady=0, padx=0, bg='green', fg='white',
                font=('Helvetica', 9, 'bold'))
titulo.grid(row=0, column=1, columnspan=5, pady=10, sticky=W+E)

variavel = StringVar()
variavel.set(" ")
variavel2 = StringVar()
variavel2.set(" ")
variavel3 = StringVar()
variavel3.set(" ")
validacao = OptionMenu(frame_2, variavel, *dias)
validacao2 = OptionMenu(frame_2, variavel2, *meses)
validacao3 = OptionMenu(frame_2, variavel3, *anos)

dia_etiqueta = Label(frame_2, text='Dia', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 8, 'bold'))
mes_etiqueta = Label(frame_2, text='Mês', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 8, 'bold'))
ano_etiqueta = Label(frame_2, text='Ano', padx=8, pady=0, bg='green', fg='white', bd = 2, relief = SUNKEN,
                     font=('Helvetica', 8, 'bold'))



dia_etiqueta.grid(row=1, column=1, pady=0, ipadx=14, ipady=0)
mes_etiqueta.grid(row=1, column=2, pady=0, ipadx=14, ipady=0)
ano_etiqueta.grid(row=1, column=3, pady=0, ipadx=14, ipady=0)
validacao.grid(row=2, column=1, pady=0)
validacao2.grid(row=2, column=2, pady=0)
validacao3.grid(row=2, column=3, pady=0)

botao_executar = Button(frame_2, text='Analisar\ncertidões', command=executa, padx=25, pady=10, bg='white', fg='green',
                        font=('Helvetica', 10, 'bold'), bd=1)
botao_renomear = Button(frame_2, text='Renomear\ncertidões', command=renomeia, padx=20, pady=10, bg='white', fg='green',
                        font=('Helvetica', 10, 'bold'), bd=1)
botao_executar.grid(row=3, column=4, padx=15, pady=50)
botao_renomear.grid(row=3, column=5, padx=0,pady=50)

titulo = Label(frame_2,  text = "SRSSU/DA/GEOF", pady=0, padx=0, bg='green', fg='white',
                font=('Helvetica', 9, 'italic'), anchor=E)
titulo.grid(row=4, column=1, columnspan=5, pady=0, sticky=W+E)

tela.mainloop()
