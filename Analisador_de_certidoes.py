from tkinter import *
from tkinter import filedialog
from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time
from pdf2image import convert_from_path
from PIL import Image
import os
import pytesseract
import re
import shutil
import PyPDF2

class Analisador:
    def __init__(self, tela):
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)

        self.frame_data = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.frame_renomear = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.menu_certidões = Menu(tela)
        self.menu_configurações = Menu(self.menu_certidões)
        self.menu_certidões.add_cascade(label='Configurações', menu=self.menu_configurações)
        self.menu_configurações.add_separator()
        self.menu_configurações.add_command(label='Caminhos', command=self.abrir_janela_caminhos)
        self.menu_configurações.add_separator()


        self.titulo = Label(self.frame_data, text='''    Indique a data limite pretendida para o próximo pagamento e    
        em seguida escolha uma das seguintes opções:    ''', pady=0, padx=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                       font=('Helvetica', 12, 'bold'))

        self.dia_etiqueta = Label(self.frame_data, text='Dia', padx=22, pady=12, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 9, 'bold'))
        self.mes_etiqueta = Label(self.frame_data, text='Mês', padx=22, pady=12, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 9, 'bold'))
        self.ano_etiqueta = Label(self.frame_data, text='Ano', padx=22, pady=12, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 9, 'bold'))

        self.variavel = StringVar()
        self.variavel.set(" ")
        self.variavel2 = StringVar()
        self.variavel2.set(" ")
        self.variavel3 = StringVar()
        self.variavel3.set(" ")
        self.dias = [' ']
        self.meses = [' ']
        self.anos = [' ']

        self.cria_calendario()

        self.validacao = OptionMenu(self.frame_data, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(self.frame_data, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(self.frame_data, self.variavel3, *self.anos)

        self.titulo_analisar = Label(self.frame_mestre, text='''Utilize esta opção para identificar quais certidões devem ser atualizadas ou
        se há requisitos a cumprir para a devida execução da análise.''', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_analisar = Button(self.frame_mestre, text='Analisar\ncertidões', command=self.executa, padx=30, pady=1, bg='green',
                                fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_transfere_arquivos = Label(self.frame_mestre, text='''Esta opção transfere as certidões para uma pasta identificada pela data
        do pagamento. Esse passo deve ser executado logo após a análise.''', pady=0, padx=0, bg='white', fg='black',
                                          font=('Helvetica', 9, 'bold'))

        self.botao_transfere_arquivos = Button(self.frame_mestre, text='Transferir\ncertidões', command=self.transfere_certidoes,
                                          padx=30, pady=1, bg='green',
                                          fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_mescla_arquivos = Label(self.frame_mestre, text='''Se já houve o pagamento e os comprovantes estão na devida pasta, esta 
        opção mescla os comprovantes com suas respectivas certidões.''', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))

        self.botao_mescla_arquivos = Button(self.frame_mestre, text=' Mesclar  \narquivos', command=self.mescla_certidoes, padx=30,
                                       pady=1, bg='green',
                                       fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.roda_pe = Label(self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
                             font=('Helvetica', 8, 'italic'), anchor=E)

        self.titulo_renomear = Label(self.frame_mestre, text='''Após atualizar as certidões, use esta opção para padronizar os nomes dos 
                        arquivos e em seguida faça nova análise para certificar que está tudo OK.''', pady=0, padx=0,
                                     bg='white', fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_procurar_arquivo = Button(self.frame_renomear, text=' Procurar\narquivo ', command=self.caminho_de_arquivo,
                                             padx=0, pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.arquivo_selecionado = 'Endereço do arquivo'
        self.caminho_do_arquivo = Label(self.frame_renomear, text=self.arquivo_selecionado, pady=0, padx=130, bg='white',fg='gray',
                                        font=('Helvetica', 9, 'bold'))

        self.botao_renomear_arquivo = Button(self.frame_renomear, text=' Renomear  \narquivo', command=self.pdf_para_jpg_para_renomear_arquivo,
                                             padx=12, pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.botao_procurar_pasta = Button(self.frame_renomear, text='   Listar   \npastas', command=self.renomeia,
                                             padx=0, pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'),
                                             bd=1)

        self.caminho_da_pasta = Label(self.frame_renomear, text='Lista de pastas', pady=0, padx=130, bg='white',
                                        fg='gray',
                                        font=('Helvetica', 9, 'bold'))

        self.botao_renomear_pasta = Button(self.frame_renomear, text=' Renomear\npastas', command=self.mescla_certidoes,
                                     padx=12,
                                     pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.botao_renomear_tudo = Button(self.frame_renomear, text='Renomear\ntodas as\ncertidões', command=self.renomeia,
                                             padx=25, pady=20, bg='green', fg='white', font=('Helvetica', 9, 'bold'),
                                             bd=1)

        self.frame_data.grid(row=0, column=1, columnspan=7, rowspan=1, pady=0, sticky=W + E)
        self.titulo.grid(row=0, column=1, columnspan=5, rowspan=1, pady=0, sticky=W + E)
        self.dia_etiqueta.grid(row=0, column=6, pady=0, ipadx=0, ipady=0)
        self.mes_etiqueta.grid(row=0, column=7, pady=0, ipadx=0, ipady=0)
        self.ano_etiqueta.grid(row=0, column=8, pady=0, ipadx=0, ipady=0)
        self.validacao.grid(row=1, column=6, pady=0)
        self.validacao2.grid(row=1, column=7, pady=0)
        self.validacao3.grid(row=1, column=8, pady=0)

        self.titulo_analisar.grid(row=1, column=1,  columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_analisar.grid(row=2, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_transfere_arquivos.grid(row=3, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_transfere_arquivos.grid(row=4, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_mescla_arquivos.grid(row=5, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_mescla_arquivos.grid(row=6, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_renomear.grid(row=7, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)

        self.frame_renomear.grid(row=8, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_procurar_arquivo.grid(row=0, column=1, padx=0, pady=0)
        self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_renomear_arquivo.grid(row=0, column=3, padx=0, pady=0)
        self.botao_procurar_pasta.grid(row=1, column=1, padx=0, pady=0)
        self.caminho_da_pasta.grid(row=1, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_renomear_pasta.grid(row=1, column=3, padx=0, pady=0)
        self.botao_renomear_tudo.grid(row=0, column=4, rowspan=2, padx=20, pady=5, ipady=8)

        self.roda_pe.grid(row=9, column=1, columnspan=10, pady=5, sticky=W+E)




    def cria_calendario(self):
        contador_dia = 1
        while contador_dia <= 31:
            if contador_dia < 10:
                self.dias.append(f"0{contador_dia}")
                contador_dia += 1
            else:
                self.dias.append(str(contador_dia))
                contador_dia += 1
        contador_mes = 1
        while contador_mes <= 12:
            if contador_mes < 10:
                self.meses.append(f"0{contador_mes}")
                contador_mes += 1
            else:
                self.meses.append(str(contador_mes))
                contador_mes += 1
        contador_anos = 2010
        while contador_anos <= 2040:
            self.anos.append(str(contador_anos))
            contador_anos += 1
        return self.dias, self.meses, self.anos

    def executa(self):
        tempo_inicial = time.time()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        obj1 = Certidao(dia, mes, ano)

        obj1.mensagem_log('Início da execução')

        obj1.analisa_referencia()
        obj1.dados_completos_dos_fornecedores()
        obj1.listar_cnpjs()

        obj1.mensagem_log_sem_horario('\nFornecedores analisados:')
        for emp in obj1.empresas:
            obj1.mensagem_log_sem_horario(f'{emp}')

        obj1.cria_diretorio()
        obj1.apaga_imagem()
        obj1.certidoes_n_encontradas()
        obj1.pdf_para_jpg()
        obj1.analisa_certidoes()

        obj1.mensagem_log_sem_horario('\nRESULTADO DA CONFERÊNCIA:')

        obj1.pega_cnpj()

        obj1.mensagem_log_sem_horario('\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n')

        for emp in obj1.empresas_a_atualizar:
            obj1.mensagem_log_sem_horario(f'{emp} - {obj1.empresas_a_atualizar[emp][0:-1]} - CNPJ: {obj1.empresas_a_atualizar[emp][-1]}\n')

        obj1.apaga_imagem()

        tempo_final = time.time()
        tempo_de_execução = int((tempo_final - tempo_inicial))
        obj1.mensagem_log(
            f'\n\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
        obj1.mensagem_log_sem_horario(
            '\n\n======================================================================================\n')

    def renomeia(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        obj1 = Certidao(dia, mes, ano)
        obj1.mensagem_log('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES INICIADO:\n')
        obj1.analisa_referencia()
        obj1.pega_fornecedores()
        obj1.apaga_imagem()
        obj1.pdf_para_jpg_renomear()
        obj1.gera_nome()
        obj1.apaga_imagem()

    def transfere_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        obj1 = Certidao(dia, mes, ano)
        obj1.analisa_referencia()
        obj1.pega_fornecedores()
        obj1.certidoes_para_pagamento()

    def mescla_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        obj1 = Certidao(dia, mes, ano)
        obj1.analisa_referencia()
        obj1.pega_fornecedores()
        obj1.merge()


    def pdf_para_jpg_para_renomear_arquivo(self):
        print('CRIANDO IMAGENS:\n')
        certidão_pdf = self.arquivo_selecionado
        pages = convert_from_path(certidão_pdf, 300, last_page=1)
        certidão_convertida_para_jpg = f'{certidão_pdf[:-4]}.jpg'
        pages[0].save(certidão_convertida_para_jpg, "JPEG")
        print('\nIMAGENS CRIADAS COM SUCESSO!')

        certidao_jpg = pytesseract.image_to_string(Image.open(certidão_convertida_para_jpg), lang='por')
        padroes = ['FGTS - CRF', 'Brasília,', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA', 'GOVERNO DO DISTRITO FEDERAL']
        valores = {'FGTS - CRF': 'FGTS', 'Brasília,': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST',
                               'MINISTÉRIO DA FAZENDA': 'UNIÃO', 'GOVERNO DO DISTRITO FEDERAL':'GDF'}
        datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',
                             'Brasília,': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                          '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)',
                             'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',
                             'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',
                             'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                             '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
        datas2 = {'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                             '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)'}

        for frase in padroes:
            if frase in certidao_jpg:
                print('Renomeando certidão')
            if frase == 'GOVERNO DO DISTRITO FEDERAL':
                try:
                    data = re.compile(datas2[frase])
                    procura = data.search(certidao_jpg)
                    datanome = procura.group()
                    separa = datanome.split('/')
                    junta = '-'.join(separa)
                except AttributeError:
                    data = re.compile(datas[frase])
                    procura = data.search(certidao_jpg)
                    datanome = procura.group()
                    separa = datanome.split('/')
                    junta = '-'.join(separa)
            else:
                data = re.compile(datas[frase])
                procura = data.search(certidao_jpg)
                datanome = procura.group()
                separa = datanome.split('/')
                junta = '-'.join(separa)
            if ':' in junta:
                retira = junta.split(':')
                volta = ' '.join(retira)
                junta = volta
        shutil.move(f'{certidão_convertida_para_jpg[0:-4]}.pdf', f'{valores[frase]} - {junta}.pdf')
        print('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES EXECUTADO COM SUCESSO!')

    def caminho_de_arquivo(self):
        self.arquivo_selecionado = filedialog.askopenfilename(initialdir='//hrg-74977/GEOF/CERTIDÕES/Certidões2')
        self.caminho_do_arquivo = Label(self.frame_renomear, text=f'...{self.arquivo_selecionado[-15:]}', pady=0, padx=0, bg='white',
                                        fg='gray', font=('Helvetica', 9, 'bold'))
        self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W+E)

    def caminho_de_pastas(self):
        pasta = filedialog.askdirectory(initialdir='C:/Users/14343258/Desktop')
        self.titulo_caminhos = Label(self.frame_de_caminhos, text='', pady=0, padx=0, bg='white', fg='black',
                       font=('Helvetica', 9, 'bold'))
        self.titulo_caminhos.grid(row=1, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)


    def abrir_janela_caminhos(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        obj1 = Certidao(dia, mes, ano)
        self.janela_de_caminhos = Toplevel()
        self.janela_de_caminhos.title('Lista de caminhos')
        self.janela_de_caminhos.resizable(False, False)
        self.frame_de_caminhos = LabelFrame(self.janela_de_caminhos, padx=0, pady=0)
        self.frame_de_caminhos.pack(padx=1, pady=1)
        self.botao_xlsx = Button(self.frame_de_caminhos, text='Fonte de dados XLSX', command=self.caminho_de_pastas,
                                 padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_xlsx = Label(self.frame_de_caminhos, text=obj1.caminho_xls, pady=0, padx=0, bg='white', fg='black',
                                  font=('Helvetica', 9, 'bold'))
        self.botao_pasta_de_certidões = Button(self.frame_de_caminhos, text='Fonte', command=self.caminho_de_pastas,
                                    padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_pasta_de_certidões = Label(self.frame_de_caminhos, text='-', pady=0, padx=0, bg='white', fg='black',
                                     font=('Helvetica', 9, 'bold'))
        self.botao_log = Button(self.frame_de_caminhos, text='Fonte', command=self.caminho_de_pastas, padx=10, pady=10,
                                bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_log = Label(self.frame_de_caminhos, text='-', pady=0, padx=0, bg='white', fg='black',
                                                font=('Helvetica', 9, 'bold'))
        self.pasta_pagamento = Button(self.frame_de_caminhos, text='Fonte', command=self.caminho_de_pastas,
                                      padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_pasta_pagamento = Label(self.frame_de_caminhos, text='-', pady=0, padx=0, bg='white',fg='black',
                                 font=('Helvetica', 9, 'bold'))
        self.certidões_para_pagamento = Button(self.frame_de_caminhos, text='Fonte', command=self.caminho_de_pastas, padx=10,
                                      pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_certidões_para_pagamento = Label(self.frame_de_caminhos, text='-', pady=0, padx=0, bg='white',
                                                      fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_xlsx.grid(row=1, column=1, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
        self.caminho_xlsx.grid(row=1, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_pasta_de_certidões.grid(row=2, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_de_certidões.grid(row=2, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_log.grid(row=3, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_log.grid(row=3, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.pasta_pagamento.grid(row=4, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_pagamento.grid(row=4, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.certidões_para_pagamento.grid(row=5, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_certidões_para_pagamento.grid(row=5, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)

tela = Tk()

objeto_tela = Analisador(tela)
tela.resizable(False, False)
tela.title('GEOF - Analisador de certidões')
#tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')
tela.config(menu=objeto_tela.menu_certidões)


tela.mainloop()
