from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time
from pdf2image import convert_from_path
from PIL import Image
import os
import pytesseract
import re
import shutil
import PyPDF2
import sqlite3

class Analisador:
    def __init__(self, tela):
        self.urls = []
        self.cria_bd()
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

        self.botao_abrir_log = Button(self.frame_data, text='Abrir log', command=self.abrir_log, padx=0,
                                     pady=0, bg='white',
                                     fg='green', font=('Helvetica', 9, 'bold'), bd=1)
        self.validacao = OptionMenu(self.frame_data, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(self.frame_data, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(self.frame_data, self.variavel3, *self.anos)

        self.titulo_analisar = Label(self.frame_mestre, text='''Utilize esta opção para identificar quais certidões devem ser atualizadas ou
        se há requisitos a cumprir para a devida execução da análise.''', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_analisar = Button(self.frame_mestre, text='Analisar\ncertidões', command=self.executa, padx=30, pady=1, bg='green',
                                fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.titulo_renomear = Label(self.frame_mestre, text='''Após atualizar as certidões, use uma das opções (renomear arquivos, pasta ou todas as certidões) para padronizar
         os nomes dos arquivos e em seguida faça nova análise para certificar que está tudo OK.''', pady=0, padx=0,
                                     bg='white', fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_procurar_arquivo = Button(self.frame_renomear, text=' Selecionar\narquivos ',
                                             command=self.caminho_de_arquivo,
                                             padx=0, pady=1, bg='white', fg='green', font=('Helvetica', 9, 'bold'),
                                             bd=1)

        self.arquivo_selecionado = 'Selecione os arquivos que deseja renomear'
        self.caminho_do_arquivo = Label(self.frame_renomear, text=self.arquivo_selecionado, pady=0, padx=50, bg='white',
                                        fg='gray',
                                        font=('Helvetica', 9, 'bold'))

        self.botao_renomear_arquivo = Button(self.frame_renomear, text=' Renomear  \narquivos',
                                             command=self.pdf_para_jpg_para_renomear_arquivo,
                                             padx=12, pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'),
                                             bd=1)

        self.botao_procurar_pasta = Button(self.frame_renomear, text='Selecionar\npasta',
                                           command=self.caminho_de_pastas,
                                           padx=0, pady=1, bg='white', fg='green', font=('Helvetica', 9, 'bold'),
                                           bd=1)
        self.pasta_selecionada = 'Selecione a pasta que deseja renomear'
        self.caminho_da_pasta = Label(self.frame_renomear, text=self.pasta_selecionada, pady=0, padx=50, bg='white',
                                      fg='gray',
                                      font=('Helvetica', 9, 'bold'))

        self.botao_renomear_pasta = Button(self.frame_renomear, text=' Renomear\npasta',
                                           command=self.pdf_para_jpg_renomear_conteudo_da_pasta,
                                           padx=12,
                                           pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.botao_renomear_tudo = Button(self.frame_renomear, text='Renomear\ntodas as\ncertidões',
                                          command=self.renomeia,
                                          padx=25, pady=20, bg='green', fg='white', font=('Helvetica', 9, 'bold'),
                                          bd=1)

        self.titulo_transfere_arquivos = Label(self.frame_mestre, text='''Esta opção transfere as certidões que validam o pagamento para uma pasta identificada pela data.
        Esse passo deve ser executado logo após a análise definitiva antes do pagamento.''', pady=0, padx=0, bg='white',
                                               fg='black',
                                               font=('Helvetica', 9, 'bold'))

        self.botao_transfere_arquivos = Button(self.frame_mestre, text='Transferir\ncertidões', command=self.transfere_certidoes,
                                          padx=30, pady=1, bg='green',
                                          fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_mescla_arquivos = Label(self.frame_mestre, text='''Após o pagamento utilize esta opção para mesclar os comprovantes
        de pagamento digitalizados com suas respectivas certidões.''', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))

        self.botao_mescla_arquivos = Button(self.frame_mestre, text=' Mesclar  \narquivos', command=self.mescla_certidoes, padx=30,
                                       pady=1, bg='green',
                                       fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.roda_pe = Label(self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
                             font=('Helvetica', 8, 'italic'), anchor=E)

        self.frame_data.grid(row=0, column=1, columnspan=7, rowspan=1, pady=0, sticky=W + E)
        self.titulo.grid(row=0, column=1, columnspan=5, rowspan=1, pady=0, sticky=W + E)
        self.dia_etiqueta.grid(row=0, column=6, pady=0, ipadx=0, ipady=0)
        self.mes_etiqueta.grid(row=0, column=7, pady=0, ipadx=0, ipady=0)
        self.ano_etiqueta.grid(row=0, column=8, pady=0, ipadx=0, ipady=0)
        self.botao_abrir_log.grid(row=1, column=1, pady=0)
        self.validacao.grid(row=1, column=6, pady=0)
        self.validacao2.grid(row=1, column=7, pady=0)
        self.validacao3.grid(row=1, column=8, pady=0)

        self.titulo_analisar.grid(row=1, column=1,  columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_analisar.grid(row=2, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_renomear.grid(row=3, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)

        self.frame_renomear.grid(row=4, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_procurar_arquivo.grid(row=0, column=1, padx=0, pady=0)
        self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_renomear_arquivo.grid(row=0, column=3, padx=0, pady=0)
        self.botao_procurar_pasta.grid(row=1, column=1, padx=0, pady=0)
        self.caminho_da_pasta.grid(row=1, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_renomear_pasta.grid(row=1, column=3, padx=0, pady=0)
        self.botao_renomear_tudo.grid(row=0, column=4, rowspan=2, padx=20, pady=5, ipady=8)


        self.titulo_transfere_arquivos.grid(row=5, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_transfere_arquivos.grid(row=6, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_mescla_arquivos.grid(row=7, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W + E)
        self.botao_mescla_arquivos.grid(row=8, column=1, columnspan=7, padx=0, pady=10)

        self.roda_pe.grid(row=9, column=1, columnspan=10, pady=5, sticky=W+E)


    def __file__(self):
        caminho_py = __file__
        caminho_do_dir = caminho_py.split('\\')
        caminho_uso = ('/').join(caminho_do_dir[0:-1])
        return caminho_uso


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
        self.botao_xlsx = Button(self.frame_de_caminhos, text='Fonte de dados XLSX', command=self.caminho_geral,
                                 padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=30)
        #nome.grid(row=0, column=1, padx=20, pady=(10, 0), sticky=W + E + S + N)

        #self.caminho_xlsx = Label(self.frame_de_caminhos, text=self.urls[0][1], pady=0, padx=0, bg='white', fg='black',
                                  #font=('Helvetica', 9, 'bold'))
        self.botao_pasta_de_certidões = Button(self.frame_de_caminhos, text='Pasta de certidões', command=self.caminho_de_pastas,
                                    padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_pasta_de_certidões = Label(self.frame_de_caminhos, text=self.urls[1][1], pady=0, padx=0, bg='white', fg='black',
                                     font=('Helvetica', 9, 'bold'))
        self.botao_log = Button(self.frame_de_caminhos, text='Pasta de logs', command=self.caminho_de_pastas, padx=10, pady=10,
                                bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_log = Label(self.frame_de_caminhos, text=self.urls[2][1], pady=0, padx=0, bg='white', fg='black',
                                                font=('Helvetica', 9, 'bold'))
        self.pasta_pagamento = Button(self.frame_de_caminhos, text='Pasta de pagamentos', command=self.caminho_de_pastas,
                                      padx=10, pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_pasta_pagamento = Label(self.frame_de_caminhos, text=self.urls[3][1], pady=0, padx=0, bg='white',fg='black',
                                 font=('Helvetica', 9, 'bold'))
        self.certidões_para_pagamento = Button(self.frame_de_caminhos, text='Certidões para pagamento', command=self.caminho_de_pastas, padx=10,
                                      pady=10, bg='green', fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        self.caminho_certidões_para_pagamento = Label(self.frame_de_caminhos, text=self.urls[4][1], pady=0, padx=0, bg='white',
                                                      fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_xlsx.grid(row=1, column=1, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
        self.caminho_xlsx.grid(row=1, column=2, padx=20)
        #self.caminho_xlsx.grid(row=1, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_pasta_de_certidões.grid(row=2, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_de_certidões.grid(row=2, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_log.grid(row=3, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_log.grid(row=3, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.pasta_pagamento.grid(row=4, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_pagamento.grid(row=4, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.certidões_para_pagamento.grid(row=5, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_certidões_para_pagamento.grid(row=5, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)


    def cria_bd(self):
        if not os.path.exists(f'{self.__file__()}/caminhos.db'):
            conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute('CREATE TABLE urls (variavel text, url text)')
            caminhos = {'caminho_xlsx': "//hrg-74977/GEOF/CERTIDÕES/Análise/atual.xlsx",
                        'pdf_dir': "//hrg-74977/GEOF/CERTIDÕES/Certidões2",
                        'caminho_de_log': '//hrg-74977/GEOF/CERTIDÕES/Logs de conferência',
                        'pasta_de_trabalho': '//hrg-74977/GEOF/HRG/PDPAS 2020/PAGAMENTO',
                        'pagamento_por_data': '//hrg-74977/GEOF/CERTIDÕES/Pagamentos'}
            for caminho in caminhos:
                direcionador.execute('INSERT INTO urls VALUES (:variavel, :url)',
                                     {"variavel": caminho, "url": caminhos[caminho]})
                conexao.commit()
            direcionador.execute("SELECT *, oid FROM urls")
            self.urls = direcionador.fetchall()
            for registro in self.urls:
                print(registro)
            conexao.close()
        else:
            print('Banco de dados localizado.')
            conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute("SELECT *, oid FROM urls")
            self.urls = direcionador.fetchall()
            for registro in self.urls:
                print(registro)
                conexao.close()


    def abrir_log(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        if not os.path.exists(f'\\\hrg-74977\\GEOF\\CERTIDÕES\\Logs de conferência\\{ano}-{mes}-{dia}.txt') or (dia, mes, ano) == (' ', ' ', ' '):
            messagebox.showerror('Me ajuda a te ajudar!',
                                 'Não existe log para a data informada.')
        else:
            os.startfile(f'\\\hrg-74977\\GEOF\\CERTIDÕES\\Logs de conferência\\{ano}-{mes}-{dia}.txt')


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

        obj1.mensagem_log('\n====================================================================================================================================\n\nInício da execução')

        obj1.analisa_referencia()
        obj1.dados_completos_dos_fornecedores()
        obj1.listar_cnpjs()
        obj1.listar_cnpjs_exceções()

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

        print(obj1.empresas)

        tempo_final = time.time()
        tempo_de_execução = int((tempo_final - tempo_inicial))
        obj1.mensagem_log(
            f'\n\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
        obj1.mensagem_log_sem_horario(
            '\n\n====================================================================================================================================\n')
        messagebox.showinfo('Analisou, miserávi!', 'Processo de análise de certidões executado com sucesso!')

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
        messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões da listagem de pagamento foram renomeadas com sucesso!')

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


    def caminho_de_arquivo(self):
        self.arquivo_selecionado = filedialog.askopenfilenames(initialdir='//hrg-74977/GEOF/CERTIDÕES/Certidões2',
                                                               filetypes=(('Arquivos pdf','*.pdf'),("Todos os arquivos", '*.*')))
        numero_de_arquivos = 'Nenhum arquivo selecionado'
        if len(self.arquivo_selecionado) > 1:
            numero_de_arquivos = 'Multiplos arquivos selecionados'
        elif len(self.arquivo_selecionado) == 1:
            numero_de_arquivos = os.path.basename(self.arquivo_selecionado[0])
        self.caminho_do_arquivo = Label(self.frame_renomear, text=numero_de_arquivos, pady=0,
                                        padx=50, bg='white', fg='gray', font=('Helvetica', 9, 'bold'))
        self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W+E)



    def pdf_para_jpg_para_renomear_arquivo(self):
        if self.arquivo_selecionado == 'Selecione os arquivos que deseja renomear' or list(self.arquivo_selecionado) == []:
            messagebox.showerror('Se não selecionar os arquivos, não vai rolar!', 'Selecione os arquivos que deseja renomear')
            print('Selecione os arquivos que deseja renomear')
        elif not os.path.exists(self.arquivo_selecionado[0]):
            print('O arquivo selecionado não existe.')
            messagebox.showerror('Esse arquivo é invenção da sua cabeça, parça!',
                                 'O arquivo selecionado não existe ou já foi renomeado!')
            self.caminho_do_arquivo = Label(self.frame_renomear, text='O arquivo selecionado não existe.', pady=0,
                                            padx=50, bg='white', fg='gray', font=('Helvetica', 9, 'bold'))
            self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            print('CRIANDO IMAGENS:\n')
            certidão_pdf = list(self.arquivo_selecionado)
            print(certidão_pdf)
            for arquivo_a_renomear in certidão_pdf:
                os.chdir(arquivo_a_renomear[0:-((arquivo_a_renomear[::-1].find('/')+1))])
                pages = convert_from_path(arquivo_a_renomear, 300, last_page=1)
                certidão_convertida_para_jpg = f'{arquivo_a_renomear[:-4]}.jpg'
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
                        os.unlink(certidão_convertida_para_jpg)
            print('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES EXECUTADO COM SUCESSO!')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões selecionadas foram renomeadas com sucesso!')


    def caminho_geral(self):
        caminho = filedialog.askopenfilename(initialdir=self.__file__(),filetypes=(('Arquivos', '*.xlsx'),
                                                                                   ("Todos os arquivos", '*.*')))
        return caminho

    def caminho_xlsx(self):
        caminho_xls = self.caminho_geral()
        conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        direcionador = conexao.cursor()
        direcionador.execute("UPDATE urls SET url = :url WHERE oid = 1", {"url": caminho_xls})
        conexao.commit()
        #conexao.close()

        #conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        #direcionador = conexao.cursor()
        direcionador.execute("SELECT *, oid FROM urls")
        self.urls = direcionador.fetchall()
        for registro in self.urls:
            print(registro)
            conexao.close()
        self.janela_de_caminhos.destroy()

    def caminho_de_pastas(self):
        pasta = 'Nenhuma pasta selecionada'
        self.pasta_selecionada = filedialog.askdirectory(initialdir='//hrg-74977/GEOF/CERTIDÕES/Certidões2')
        if os.path.isdir(self.pasta_selecionada) and self.pasta_selecionada != '//hrg-74977/GEOF/CERTIDÕES/Certidões2':
            pasta = self.pasta_selecionada
            self.caminho_da_pasta = Label(self.frame_renomear, text=os.path.basename(pasta), pady=0, padx=0, bg='white', fg='gray',
                       font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            self.caminho_da_pasta = Label(self.frame_renomear, text=pasta, pady=0, padx=0, bg='white',
                                          fg='gray',
                                          font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)


    def apaga_imagens_da_pasta(self):
            os.chdir(self.pasta_selecionada)
            for arquivo in os.listdir(self.pasta_selecionada):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pasta_selecionada}/{arquivo}')

    def pdf_para_jpg_renomear_conteudo_da_pasta(self):
        if self.pasta_selecionada == 'Selecione a pasta que deseja renomear' or self.pasta_selecionada =='':
            messagebox.showerror('Se não selecionar a pasta, não vai rolar!',
                                 'Selecione uma pasta que contenha certidões que precisam ser renomeadas.')
            print('nenhuma pasta selecionada')
            self.caminho_da_pasta = Label(self.frame_renomear, text='Nenhuma pasta selecionada', pady=0, padx=0, bg='white',
                                          fg='gray',
                                          font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            print('CRIANDO IMAGENS:\n')
            os.chdir(self.pasta_selecionada)
            self.apaga_imagens_da_pasta()
            for pdf_file in os.listdir(self.pasta_selecionada):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pasta_selecionada}/Merge'):
                        os.makedirs(f'{self.pasta_selecionada}/Merge')
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Merge/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Merge/{pdf_file}')

                elif pdf_file.endswith(".pdf"):
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

            print('\nRENOMEANDO CERTIDÕES:\n\n')
            os.chdir(f'{self.pasta_selecionada}')
            origem = f'{self.pasta_selecionada}'
            for imagem in os.listdir(origem):
                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(Image.open(f'{origem}/{imagem}'), lang='por')
                    padroes = ['FGTS - CRF', 'Brasília,', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA',
                                   'GOVERNO DO DISTRITO FEDERAL']
                    valores = {'FGTS - CRF': 'FGTS', 'Brasília,': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST',
                                   'MINISTÉRIO DA FAZENDA': 'UNIÃO', 'GOVERNO DO DISTRITO FEDERAL': 'GDF'}
                    datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'Brasília,': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                              '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)',
                                 'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                                '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                                '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
                    datas2 = {
                            'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                           '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                           '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)'}
                    for frase in padroes:
                        if frase in certidao:
                            if frase == 'GOVERNO DO DISTRITO FEDERAL':
                                try:
                                    data = re.compile(datas2[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                                except AttributeError:
                                    data = re.compile(datas[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                            else:
                                data = re.compile(datas[frase])
                                procura = data.search(certidao)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                            if ':' in junta:
                                retira = junta.split(':')
                                volta = ' '.join(retira)
                                junta = volta
                            shutil.move(f'{origem}/{imagem[0:-4]}.pdf', f'{valores[frase]} - {junta}.pdf')
            self.apaga_imagens_da_pasta()
            print('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES EXECUTADO COM SUCESSO!')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões da pasta selecionada foram renomeadas com sucesso!')




tela = Tk()

objeto_tela = Analisador(tela)
tela.resizable(0, 0)
tela.title('GEOF - Analisador de certidões')
#tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')
tela.config(menu=objeto_tela.menu_certidões)


tela.mainloop()
