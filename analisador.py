from conexao import Conexao
from certidão import Certidao
from união import Uniao
from tst import Tst
from fgts import Fgts
from gdf import Gdf

from tkinter import *
from tkinter import filedialog
from pdf2image import convert_from_path
from PIL import Image
import os
import pytesseract
import re
import time
import datetime
import shutil
import PyPDF2
from tkinter import messagebox
import sqlite3


class Analisador(Certidao):
    opções = [
        'Selecione uma opção', 'Renomear arquivos',
        'Renomear todos os arquivos de uma pasta', 'Renomear todas as certidões da lista de pagamento']

    def __init__(self, tela):
        #self.cria_pastas_de_trabalho()
        #self.cria_bd()
        #self.configura_bd()
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


        self.titulo = Label(self.frame_data, text='    Indique a data limite pretendida para o próximo pagamento e em seguida escolha uma das seguintes opções:    ',
                            pady=0, padx=0, bg='green', fg='white', bd=2, relief=SUNKEN,font=('Helvetica', 10, 'bold'))

        self.dia_etiqueta = Label(self.frame_data, text='Dia', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))
        self.mes_etiqueta = Label(self.frame_data, text='Mês', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))
        self.ano_etiqueta = Label(self.frame_data, text='Ano', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))

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
        self.validacao1 = OptionMenu(self.frame_data, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(self.frame_data, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(self.frame_data, self.variavel3, *self.anos)

        self.titulo_analisar = Label(self.frame_mestre, text='Utilize esta opção para identificar quais certidões devem ser atualizadas ou se há requisitos a cumprir para a devida execução da análise.', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_analisar = Button(self.frame_mestre, text='Analisar\ncertidões', command=self.executa, padx=30,
                                     pady=1, bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.titulo_renomear = Label(self.frame_mestre, text='Após atualizar as certidões, selecione uma das opções para padronizar os nomes dos\narquivos e em seguida faça nova análise para certificar que está tudo OK.', pady=0, padx=0,
                                     bg='white', fg='black', font=('Helvetica', 9, 'bold'))

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set("Selecione uma opção")
        self.validacao = OptionMenu(self.frame_mestre, self.variavel_de_opções, *Analisador.opções)

        self.arquivo_selecionado = 'Selecione os arquivos que deseja renomear'
        self.pasta_selecionada = 'Selecione a pasta que deseja renomear'


        self.botao_renomear_tudo = Button(self.frame_mestre, text='Renomear\ncertidões',
                                          command=self.selecionador_de_opções, padx=30, pady=1, bg='green', fg='white',
                                          font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_transfere_arquivos = Label(self.frame_mestre, text='Esta opção transfere as certidões que validam o pagamento para uma pasta identificada pela data.\nEsse passo deve ser executado logo após a análise definitiva antes do pagamento.', pady=0, padx=0, bg='white',
                                               fg='black',
                                               font=('Helvetica', 9, 'bold'))

        self.botao_transfere_arquivos = Button(self.frame_mestre, text='Transferir\ncertidões', command=self.transfere_certidoes,
                                          padx=30, pady=1, bg='green',
                                          fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_mescla_arquivos = Label(self.frame_mestre, text='Após o pagamento utilize esta opção para mesclar os comprovantes de pagamento digitalizados com suas respectivas certidões.', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))

        self.botao_mescla_arquivos = Button(self.frame_mestre, text=' Mesclar  \narquivos', command=self.mescla_certidoes, padx=30,
                                       pady=1, bg='green',
                                       fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.roda_pe = Label(self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
                             font=('Helvetica', 8, 'italic'), anchor=E)

        self.frame_data.grid(row=0, column=1, columnspan=7, rowspan=1, pady=0, sticky=W+E)
        self.titulo.grid(row=0, column=1, columnspan=5, rowspan=1, pady=0, sticky=W+E)
        self.dia_etiqueta.grid(row=0, column=6, pady=0, ipadx=0, ipady=0)
        self.mes_etiqueta.grid(row=0, column=7, pady=0, ipadx=0, ipady=0)
        self.ano_etiqueta.grid(row=0, column=8, pady=0, ipadx=0, ipady=0)
        self.botao_abrir_log.grid(row=1, column=1, pady=0)
        self.validacao1.grid(row=1, column=6, pady=0)
        self.validacao2.grid(row=1, column=7, pady=0)
        self.validacao3.grid(row=1, column=8, pady=0)

        self.titulo_analisar.grid(row=1, column=1,  columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_analisar.grid(row=2, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_renomear.grid(row=3, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.validacao.grid(row=4, column=1, columnspan=7, padx=0, pady=10)
        self.botao_renomear_tudo.grid(row=5, column=1, columnspan=7, padx=0, pady=10)

        self.titulo_transfere_arquivos.grid(row=7, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_transfere_arquivos.grid(row=8, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_mescla_arquivos.grid(row=9, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_mescla_arquivos.grid(row=10, column=1, columnspan=7, padx=0, pady=10)

        self.roda_pe.grid(row=11, column=1, columnspan=10, pady=5, sticky=W+E)


    def abrir_janela_caminhos(self):
        self.janela_de_caminhos = Toplevel()
        urls = self.consulta_urls()
        print(f'Essa é a consulta de urls {urls}')
        self.janela_de_caminhos.title('Lista de caminhos')
        self.janela_de_caminhos.resizable(False, False)
        self.frame_de_caminhos = LabelFrame(self.janela_de_caminhos, padx=0, pady=0)
        self.frame_de_caminhos.pack(padx=1, pady=1)

        self.criar_estrutura = Label(self.frame_de_caminhos, text='Se deseja criar toda a estrutura de pastas de trabalho\n'
                                                             'necessárias para o correto funcionamento do programa na\n'
                                                             'pasta que contém o arquivo principal clique em "Criar estrura\n", '
                                                             'caso contrário selecione manualmente cada caminho abaixo.\n')

        self.botão_criar_estrutura = Button(self.frame_de_caminhos, text='Criar estrutura', command=self.cria_pastas_de_trabalho,
                                 padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)

        self.botao_xlsx = Button(self.frame_de_caminhos, text='Fonte de\ndados XLSX', command=self.altera_caminho_xlsl,
                                 padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=70)
        self.botao_pasta_de_certidões = Button(self.frame_de_caminhos, text='Pasta de\ncertidões', command=self.altera_caminho_pasta_de_certidões,
                                               padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_pasta_de_certidões = Entry(self.frame_de_caminhos, width=70)
        self.botao_log = Button(self.frame_de_caminhos, text='Pasta de\nlogs', command=self.altera_caminho_log, padx=0, pady=0,
                                bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_log = Entry(self.frame_de_caminhos, width=70)
        self.pasta_pagamento = Button(self.frame_de_caminhos, text='Comprovantes\nde pagamentos', command=self.altera_caminho_pasta_pagamento,
                                      padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_pasta_pagamento = Entry(self.frame_de_caminhos, width=70)
        self.certidões_para_pagamento = Button(self.frame_de_caminhos, text='Certidões para\npagamento', command=self.altera_caminho_certidões_para_pagamento, padx=0,
                                      pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_certidões_para_pagamento = Entry(self.frame_de_caminhos, width=70)

        self.gravar_alterações = Button(self.frame_de_caminhos, text='Gravar alterações',
                                               command=self.atualizar_xlsx, padx=10,
                                               pady=10, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)



        self.botão_criar_estrutura.grid(row=0, column=1, columnspan=1, padx=15, pady=10, ipadx=5, ipady=13, sticky=W + E)
        self.criar_estrutura.grid(row=0, column=2, padx=20)
        self.botao_xlsx.grid(row=1, column=1, columnspan=1, padx = 15, pady=10, ipadx=5, ipady=13, sticky=W+E)
        urls = self.consulta_urls()

        self.caminho_xlsx.insert(0, urls[0][1])

        self.caminho_xlsx.grid(row=1, column=2, padx=20)
        self.botao_pasta_de_certidões.grid(row=2, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_de_certidões.insert(0, urls[1][1])
        self.caminho_pasta_de_certidões.grid(row=2, column=2, padx=20)
        self.botao_log.grid(row=3, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_log.insert(0, urls[2][1])
        self.caminho_log.grid(row=3, column=2, padx=20)
        self.certidões_para_pagamento.grid(row=4, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_certidões_para_pagamento.insert(0, urls[4][1])
        self.caminho_certidões_para_pagamento.grid(row=4, column=2, padx=20)
        self.pasta_pagamento.grid(row=5, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_pagamento.insert(0, urls[3][1])
        self.caminho_pasta_pagamento.grid(row=5, column=2, padx=20)
        self.gravar_alterações.grid(row=6, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13)


    def criar_estrutura(self):
        self.cria_pastas_de_trabalho()
        self.cria_bd()
        self.configura_bd()

    def altera_caminho_xlsl(self):
        caminho = filedialog.askopenfilename(initialdir=self.caminho_do_arquivo(),filetypes=(('Arquivos', '*.xlsx'),
                                                                                   ("Todos os arquivos", '*.*')))
        self.caminho_xlsx.destroy()
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=30)
        self.caminho_xlsx.insert(0, caminho)
        self.caminho_xlsx.grid(row=1, column=2, padx=20)

    def altera_caminho_pasta_de_certidões(self):
        caminho = filedialog.askdirectory(initialdir=self.caminho_do_arquivo())
        self.caminho_pasta_de_certidões.destroy()
        self.caminho_pasta_de_certidões = Entry(self.frame_de_caminhos, width=30)
        self.caminho_pasta_de_certidões.insert(0, caminho)
        self.caminho_pasta_de_certidões.grid(row=2, column=2, padx=20)

    def altera_caminho_log(self):
        caminho = filedialog.askdirectory(initialdir=self.caminho_do_arquivo())
        self.caminho_log.destroy()
        self.caminho_log = Entry(self.frame_de_caminhos, width=30)
        self.caminho_log.insert(0, caminho)
        self.caminho_log.grid(row=3, column=2, padx=20)

    def altera_caminho_pasta_pagamento(self):
        caminho = filedialog.askdirectory(initialdir=self.caminho_do_arquivo())
        self.caminho_pasta_pagamento.destroy()
        self.caminho_pasta_pagamento = Entry(self.frame_de_caminhos, width=30)
        self.caminho_pasta_pagamento.insert(0, caminho)
        self.caminho_pasta_pagamento.grid(row=4, column=2, padx=20)

    def altera_caminho_certidões_para_pagamento(self):
        caminho = filedialog.askdirectory(initialdir=self.caminho_do_arquivo())
        self.caminho_certidões_para_pagamento.destroy()
        self.caminho_certidões_para_pagamento = Entry(self.frame_de_caminhos, width=30)
        self.caminho_certidões_para_pagamento.insert(0, caminho)
        self.caminho_certidões_para_pagamento.grid(row=5, column=2, padx=20)

    def atualizar_xlsx(self):
        resposta = messagebox.askyesno('Vc sabe o que está fazendo?','Tem certeza que deseja alterar a configuração dos caminhos de pastas e arquivos?')
        if resposta == True:
            conexao = sqlite3.connect(f'{self.caminho_do_arquivo()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute("UPDATE urls SET url = :caminho_xlsx WHERE oid = 1",
                                 {"caminho_xlsx": self.caminho_xlsx.get()})
            direcionador.execute("UPDATE urls SET url = :pasta_de_certidões WHERE oid = 2",
                                 {"pasta_de_certidões": self.caminho_pasta_de_certidões.get()})
            direcionador.execute("UPDATE urls SET url = :caminho_de_log WHERE oid = 3",
                                 {"caminho_de_log": self.caminho_log.get()})
            direcionador.execute("UPDATE urls SET url = :comprovantes_de_pagamento WHERE oid = 4",
                {"comprovantes_de_pagamento": self.caminho_pasta_pagamento.get()})
            direcionador.execute("UPDATE urls SET url = :certidões_para_pagamento WHERE oid = 5",
                                 {"certidões_para_pagamento": self.caminho_certidões_para_pagamento.get()})
            conexao.commit()
            conexao.close()
            self.janela_de_caminhos.destroy()
            print('\nOs caminhos para pastas e arquivos utilizados pelo sistema foram atualizados.\n')
            messagebox.showinfo('Fez porque quis!',"Caminhos para pastas e arquivos utilizados pelo sistema atualizados com sucesso!")
        else:
            self.janela_de_caminhos.destroy()

    def abrir_log(self):
        urls = self.consulta_urls()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        if not os.path.exists(f'{urls[2][1]}/{ano}-{mes}-{dia}.txt') or (dia, mes, ano) == (' ', ' ', ' '):
            messagebox.showerror('Me ajuda a te ajudar!',
                                 'Não existe log para a data informada.')
        else:
            caminho = f'{urls[2][1]}/{ano}-{mes}-{dia}.txt'
            novo_caminho = caminho.replace('/', '\\')
            os.startfile(novo_caminho)


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

    def checa_urls(self):
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]):
            messagebox.showerror('Sumiu!!!',
                                 'O arquivo xlsx selecionado como fonte foi apagado, removido ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Fonte de dados XLSX e '
                                 'selecione um arquivo xlsx que atenda aos critérios necessários '
                                 'para o processamento.')
        elif not os.path.exists(urls[1][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte para certidões foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Pasta de certidões e '
                                 'selecione uma pasta que contenha as certidões que devem ser analisadas.')
        elif not os.path.exists(urls[2][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte e destino para logs foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Pasta de logs e '
                                 'selecione uma pasta onde os logs serão criados.')
        elif not os.path.exists(urls[4][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como destino de cetidões para pagamento foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Cetidões para pagamento e '
                                 'selecione uma pasta para direcionar as certidões do pagamento.')
        elif not os.path.exists(urls[3][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte de comprovantes de pagamento foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Comprovantes de pagamento e '
                                 'selecione uma pasta que contenha os comprovantes de pagamento.')

    def executa(self):
        tempo_inicial = time.time()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]) or not os.path.exists(urls[1][1]) or not os.path.exists(urls[2][1])\
                or not os.path.exists(urls[4][1]) or not os.path.exists(urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            objUniao = Uniao(dia, mes, ano)
            objTst = Tst(dia, mes, ano)
            objFgts = Fgts(dia, mes, ano)
            objGdf = Gdf(dia, mes, ano)
            lista_de_objetos = [objUniao, objTst, objFgts, objGdf]


            obj1.mensagem_de_log_completa('\n====================================================================================================================================\n\nInício da execução', obj1.caminho_de_log)

            obj1.analisa_referencia()
            obj1.dados_completos_dos_fornecedores()
            obj1.listar_cnpjs()
            obj1.listar_cnpjs_exceções()

            obj1.mensagem_de_log_simples('\nFornecedores analisados:', obj1.caminho_de_log)
            for emp in obj1.empresas:
                obj1.mensagem_de_log_simples(f'{emp}', obj1.caminho_de_log)

            obj1.cria_diretorio()
            obj1.apaga_imagem()
            obj1.certidoes_n_encontradas()
            obj1.pdf_para_jpg()
            obj1.analisa_certidoes(lista_de_objetos)

            obj1.mensagem_de_log_simples('\nRESULTADO DA CONFERÊNCIA:', obj1.caminho_de_log)

            obj1.pega_cnpj()

            obj1.mensagem_de_log_simples('\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n', obj1.caminho_de_log)

            for emp in obj1.empresas_a_atualizar:
                obj1.mensagem_de_log_simples(f'{emp} - {obj1.empresas_a_atualizar[emp][0:-1]} - CNPJ: {obj1.empresas_a_atualizar[emp][-1]}\n', obj1.caminho_de_log)

            obj1.apaga_imagem()

            tempo_final = time.time()
            tempo_de_execução = int((tempo_final - tempo_inicial))
            obj1.mensagem_de_log_completa(
                f'\n\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.', obj1.caminho_de_log)
            obj1.mensagem_de_log_simples(
                '\n\n====================================================================================================================================\n', obj1.caminho_de_log)
            messagebox.showinfo('Analisou, miserávi!', 'Processo de análise de certidões executado com sucesso!')

    def selecionador_de_opções(self):
        print(self.variavel_de_opções.get())
        if self.variavel_de_opções.get() == 'Renomear arquivos':
            self.pdf_para_jpg_para_renomear_arquivo()
        elif self.variavel_de_opções.get() == 'Renomear todos os arquivos de uma pasta':
            self.pdf_para_jpg_renomear_conteudo_da_pasta()
        elif self.variavel_de_opções.get() == 'Renomear todas as certidões da lista de pagamento':
            self.renomeia()
        elif self.variavel_de_opções.get() == 'Selecione uma opção':
            messagebox.showwarning('Tem que escolher, fi!', 'Nenhuma opção selecionada!')

    def renomeia(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]) or not os.path.exists(urls[1][1]) or not os.path.exists(urls[2][1])\
                or not os.path.exists(urls[4][1]) or not os.path.exists(urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.mensagem_de_log_completa('\nProcesso de renomeação de certidões iniciado:\n')
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
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]) or not os.path.exists(urls[1][1]) or not os.path.exists(
                urls[2][1]) \
                or not os.path.exists(urls[4][1]) or not os.path.exists(urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.cria_certidoes_para_pagamento()

    def mescla_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]) or not os.path.exists(urls[1][1]) or not os.path.exists(
                urls[2][1]) \
                or not os.path.exists(urls[4][1]) or not os.path.exists(urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.merge()

    def caminho_de_arquivo(self):
        self.arquivo_selecionado = filedialog.askopenfilenames(initialdir=f'{self.caminho_do_arquivo()}/Certidões',
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
        self.arquivo_selecionado = filedialog.askopenfilenames(initialdir=f'{self.caminho_do_arquivo()}/Certidões',
                                                               filetypes=(
                                                               ('Arquivos pdf', '*.pdf'), ("Todos os arquivos", '*.*')))
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
            print('==================================================================================================\n'
                  'Criando imagem:\n')
            certidão_pdf = list(self.arquivo_selecionado)
            print(f'Arquivo que está sendo renomeado: {certidão_pdf}\n')
            for arquivo_a_renomear in certidão_pdf:
                os.chdir(arquivo_a_renomear[0:-((arquivo_a_renomear[::-1].find('/')+1))])
                pages = convert_from_path(arquivo_a_renomear, 300, last_page=1)
                certidão_convertida_para_jpg = f'{arquivo_a_renomear[:-4]}.jpg'
                pages[0].save(certidão_convertida_para_jpg, "JPEG")
                print('\nImagem criada com sucesso!\n')

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

                print('Renomeando certidão')
                for frase in padroes:
                    if frase in certidao_jpg:
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
            print('\nProcesso de renomeação de certidão executado com sucesso!\n'
                  '===================================================================================================')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões selecionadas foram renomeadas com sucesso!')



    def caminho_de_pastas(self):
        pasta = 'Nenhuma pasta selecionada'
        self.pasta_selecionada = filedialog.askdirectory(initialdir=f'{self.caminho_do_arquivo()}/Certidões')
        if os.path.isdir(self.pasta_selecionada) and self.pasta_selecionada != f'{self.caminho_do_arquivo()}/Certidões':
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
        self.pasta_selecionada = filedialog.askdirectory(initialdir=f'{self.caminho_do_arquivo()}/Certidões')
        if self.pasta_selecionada == 'Selecione a pasta que deseja renomear' or self.pasta_selecionada =='':
            messagebox.showerror('Se não selecionar a pasta, não vai rolar!',
                                 'Selecione uma pasta que contenha certidões que precisam ser renomeadas.')
            print('nenhuma pasta selecionada')
            self.caminho_da_pasta = Label(self.frame_renomear, text='Nenhuma pasta selecionada', pady=0, padx=0, bg='white',
                                          fg='gray',
                                          font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            print('==================================================================================================\n'
                  'Criando imagens:\n')
            os.chdir(self.pasta_selecionada)
            self.apaga_imagens_da_pasta()
            for pdf_file in os.listdir(self.pasta_selecionada):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pasta_selecionada}/Mesclados'):
                        os.makedirs(f'{self.pasta_selecionada}/Mesclados')
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Mesclados/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Mesclados/{pdf_file}')

                elif pdf_file.endswith(".pdf"):
                    print(pdf_file[:-4])
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

            print(f'\nRenomeando certidões da pasta {self.pasta_selecionada}:\n\n')
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
                            print(imagem.split()[0])
            self.apaga_imagens_da_pasta()
            print('\nProcesso de renomeação de certidões executado com sucesso!')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões da pasta selecionada foram renomeadas com sucesso!')

if __name__ == '__main__':
    tela = Tk()

    objeto_tela = Analisador(tela)
    tela.resizable(False, False)
    tela.title('GEOF - Analisador de certidões')
    #tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')
    tela.config(menu=objeto_tela.menu_certidões)

    tela.mainloop()
