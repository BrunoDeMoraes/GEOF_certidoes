import os
import pytesseract
import shutil
import sqlite3
import time
from pdf2image import convert_from_path
from PIL import Image
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

from certidão import Certidao
from constantes import ANALISADOS
from constantes import ARQUIVO_INEXISTENTE
from constantes import ATUALIZAR_CAMINHOS
from constantes import ARQUIVOS_NAO_SELECIONADOS
from constantes import CAMINHOS_ATUALIZADOS
from constantes import CHECA_URL_0, CHECA_URL_1, CHECA_URL_2
from constantes import CHECA_URL_3, CHECA_URL_4
from constantes import CONFERENCIA
from constantes import CRIANDO_IMAGENS
from constantes import IDENTIFICADOR_DE_CERTIDAO
from constantes import IDENTIFICADOR_DE_VALIDADE, IDENTIFICADOR_DE_VALIDADE_2
from constantes import IDENTIFICADOR_TRADUZIDO
from constantes import INICIO_DA_EXECUCAO
from constantes import LINHA_FINAL
from constantes import LOG_INEXISTENTE
from constantes import OPCOES_DE_RENOMEACAO
from constantes import PENDENCIAS
from constantes import RENOMEACAO_EXECUTADA
from constantes import TEXTO_ANALISAR, TEXTO_CRIA_ESTRUTURA
from constantes import TEXTO_MESCLA_ARQUIVOS, TEXTO_PRINCIPAL
from constantes import TEXTO_RENOMEAR, TEXTO_TRANSFERE_ARQUIVOS
from fgts import Fgts
from gdf import Gdf
from tst import Tst
from união import Uniao


class Analisador(Certidao):
    def __init__(self, tela):
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)
        self.frame_data = LabelFrame(self.frame_mestre, padx=0, pady=0)
        self.frame_renomear = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.menu_certidões = Menu(tela)
        self.menu_configurações = Menu(self.menu_certidões)
        self.menu_certidões.add_cascade(
            label='Configurações', menu=self.menu_configurações)
        self.menu_configurações.add_separator()
        self.menu_configurações.add_command(
            label='Caminhos', command=self.abrir_janela_caminhos)
        self.menu_configurações.add_separator()

        self.titulo = Label(
            self.frame_data, text=TEXTO_PRINCIPAL, pady=0, padx=0, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 10, 'bold')
        )

        self.dia_etiqueta = Label(
            self.frame_data, text='Dia', padx=22, pady=0, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 10, 'bold')
        )
        self.mes_etiqueta = Label(
            self.frame_data, text='Mês', padx=22, pady=0, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 10, 'bold')
        )
        self.ano_etiqueta = Label(
            self.frame_data, text='Ano', padx=22, pady=0, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Helvetica', 10, 'bold')
        )

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

        self.botao_abrir_log = Button(
            self.frame_data, text='Abrir log', command=self.abrir_log, padx=0,
            pady=0, bg='white', fg='green', font=('Helvetica', 9, 'bold'),
            bd=1
        )

        self.validacao1 = OptionMenu(
            self.frame_data, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(
            self.frame_data, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(
            self.frame_data, self.variavel3, *self.anos)

        self.titulo_analisar = Label(
            self.frame_mestre, text=TEXTO_ANALISAR, pady=0, padx=0,
            bg='white', fg='black', font=('Helvetica', 9, 'bold')
        )

        self.botao_analisar = Button(
            self.frame_mestre, text='Analisar\ncertidões',
            command=self.executa, padx=30, pady=1, bg='green',
            fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.titulo_renomear = Label(
            self.frame_mestre, text=TEXTO_RENOMEAR, pady=0, padx=0,
            bg='white', fg='black', font=('Helvetica', 9, 'bold')
        )

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set("Selecione uma opção")
        self.validacao = OptionMenu(
            self.frame_mestre, self.variavel_de_opções, *OPCOES_DE_RENOMEACAO
        )

        self.botao_renomear_tudo = Button(
            self.frame_mestre, text='Renomear\ncertidões',
            command=self.selecionador_de_opções, padx=30, pady=1,
            bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.titulo_transfere_arquivos = Label(
            self.frame_mestre, text=TEXTO_TRANSFERE_ARQUIVOS,
            pady=0, padx=0, bg='white', fg='black',
            font=('Helvetica', 9, 'bold')
        )

        self.botao_transfere_arquivos = Button(
            self.frame_mestre, text='Transferir\ncertidões',
            command=self.transfere_certidoes, padx=30, pady=1,
            bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.titulo_mescla_arquivos = Label(
            self.frame_mestre, text=TEXTO_MESCLA_ARQUIVOS, pady=0, padx=0,
            bg='white', fg='black', font=('Helvetica', 9, 'bold')
        )

        self.botao_mescla_arquivos = Button(
            self.frame_mestre, text='Mesclar\narquivos',
            command=self.mescla_certidoes, padx=30, pady=1,
            bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.roda_pe = Label(
            self.frame_mestre, text="SRSSU/DA/GEOF    ", pady=0, padx=0,
            bg='green', fg='white', font=('Helvetica', 8, 'italic'), anchor=E
        )

        self.frame_data.grid(
            row=0, column=1, columnspan=7, rowspan=1, pady=0, sticky=W+E
        )
        self.titulo.grid(
            row=0, column=1, columnspan=5, rowspan=1, pady=0, sticky=W+E
        )
        self.dia_etiqueta.grid(row=0, column=6, pady=0, ipadx=0, ipady=0)
        self.mes_etiqueta.grid(row=0, column=7, pady=0, ipadx=0, ipady=0)
        self.ano_etiqueta.grid(row=0, column=8, pady=0, ipadx=0, ipady=0)

        self.botao_abrir_log.grid(row=1, column=1, pady=0)
        self.validacao1.grid(row=1, column=6, pady=0)
        self.validacao2.grid(row=1, column=7, pady=0)
        self.validacao3.grid(row=1, column=8, pady=0)
        self.titulo_analisar.grid(
            row=1, column=1, columnspan=7, padx=0,
            pady=0, ipadx=0, ipady=8, sticky=W+E
        )

        self.botao_analisar.grid(
            row=2, column=1, columnspan=7, padx=0, pady=10
        )

        self.titulo_renomear.grid(
            row=3, column=1, columnspan=7, padx=0,
            pady=0, ipadx=0, ipady=8, sticky=W+E
        )

        self.validacao.grid(row=4, column=1, columnspan=7, padx=0, pady=10)

        self.botao_renomear_tudo.grid(
            row=5, column=1, columnspan=7, padx=0, pady=10
        )

        self.titulo_transfere_arquivos.grid(
            row=7, column=1, columnspan=7, padx=0,
            pady=0, ipadx=0, ipady=8, sticky=W+E
        )

        self.botao_transfere_arquivos.grid(
            row=8, column=1, columnspan=7, padx=0, pady=10
        )

        self.titulo_mescla_arquivos.grid(
            row=9, column=1, columnspan=7, padx=0,
            pady=0, ipadx=0, ipady=8, sticky=W+E
        )

        self.botao_mescla_arquivos.grid(
            row=10, column=1, columnspan=7, padx=0, pady=10
        )

        self.roda_pe.grid(row=11, column=1, columnspan=10, pady=5, sticky=W+E)

    def abrir_janela_caminhos(self):
        self.janela_de_caminhos = Toplevel()
        self.urls = self.consulta_urls()
        self.janela_de_caminhos.title('Lista de caminhos')
        self.janela_de_caminhos.resizable(False, False)
        self.frame_de_caminhos = LabelFrame(
            self.janela_de_caminhos, padx=0, pady=0
        )
        self.frame_de_caminhos.pack(padx=1, pady=1)

        self.criar_estrutura = Label(
            self.frame_de_caminhos, text=TEXTO_CRIA_ESTRUTURA,
            bg='white', fg='green', font=('Helvetica', 10, 'bold')
        )

        self.botão_criar_estrutura = Button(
            self.frame_de_caminhos, text='Criar estrutura',
            command=self.cria_pastas_de_trabalho, padx=0, pady=0, bg='green',
            fg='white', font=('Helvetica', 10, 'bold'), bd=1
        )

        self.botao_xlsx = Button(
            self.frame_de_caminhos, text='Fonte de\ndados XLSX',
            command=lambda: self.altera_caminho(self.caminho_xlsx, True),
            padx=0, pady=0, bg='green', fg='white',
            font=('Helvetica', 8, 'bold'), bd=1
        )
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=70)

        self.botao_pasta_de_certidões = Button(
            self.frame_de_caminhos, text='Pasta de\ncertidões',
            command=lambda: (
                self.altera_caminho(self.caminho_pasta_de_certidões)),
            padx=0, pady=0, bg='green', fg='white',
            font=('Helvetica', 8, 'bold'), bd=1
        )
        self.caminho_pasta_de_certidões = Entry(
            self.frame_de_caminhos, width=70
        )

        self.botao_log = Button(
            self.frame_de_caminhos, text='Pasta de\nlogs',
            command=lambda: self.altera_caminho(self.caminho_log), padx=0,
            pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'),
            bd=1
        )
        self.caminho_log = Entry(self.frame_de_caminhos, width=70)

        self.pasta_pagamento = Button(
            self.frame_de_caminhos, text='Comprovantes\nde pagamentos',
            command=lambda: self.altera_caminho(self.caminho_pasta_pagamento),
            padx=0, pady=0, bg='green', fg='white',
            font=('Helvetica', 8, 'bold'), bd=1
        )
        self.caminho_pasta_pagamento = Entry(self.frame_de_caminhos, width=70)

        self.certidões_para_pagamento = Button(
            self.frame_de_caminhos, text='Certidões para\npagamento',
            command=lambda: (
                self.altera_caminho(self.caminho_certidões_para_pagamento)),
            padx=0, pady=0, bg='green', fg='white',
            font=('Helvetica', 8, 'bold'), bd=1
        )
        self.caminho_certidões_para_pagamento = Entry(
            self.frame_de_caminhos, width=70
        )

        self.gravar_alterações = Button(
            self.frame_de_caminhos, text='Gravar alterações',
            command=self.atualizar_caminhos, padx=10, pady=10, bg='green',
            fg='white', font=('Helvetica', 8, 'bold'), bd=1
        )

        self.botão_criar_estrutura.grid(
            row=0, column=1, columnspan=1, padx=15, pady=10, ipadx=5,
            ipady=13, sticky=W+E
        )

        self.criar_estrutura.grid(row=0, column=2, padx=20, pady=10)

        self.botao_xlsx.grid(
            row=1, column=1, columnspan=1, padx=15, pady=10, ipadx=5,
            ipady=13, sticky=W+E
        )
        self.caminho_xlsx.insert(0, self.urls[0][1])
        self.caminho_xlsx.grid(row=1, column=2, padx=20)

        self.botao_pasta_de_certidões.grid(
            row=2, column=1, columnspan=1, padx=15, pady=10, ipadx=10,
            ipady=13, sticky=W+E
        )

        self.caminho_pasta_de_certidões.insert(0, self.urls[1][1])
        self.caminho_pasta_de_certidões.grid(row=2, column=2, padx=20)

        self.botao_log.grid(
            row=3, column=1, columnspan=1, padx=15, pady=10, ipadx=10,
            ipady=13, sticky=W+E
        )

        self.caminho_log.insert(0, self.urls[2][1])
        self.caminho_log.grid(row=3, column=2, padx=20)

        self.certidões_para_pagamento.grid(
            row=4, column=1, columnspan=1, padx=15, pady=10, ipadx=10,
            ipady=13, sticky=W+E
        )

        self.caminho_certidões_para_pagamento.insert(0, self.urls[4][1])
        self.caminho_certidões_para_pagamento.grid(row=4, column=2, padx=20)

        self.pasta_pagamento.grid(
            row=5, column=1, columnspan=1, padx=15, pady=10, ipadx=10,
            ipady=13, sticky=W+E
        )

        self.caminho_pasta_pagamento.insert(0, self.urls[3][1])
        self.caminho_pasta_pagamento.grid(row=5, column=2, padx=20)

        self.gravar_alterações.grid(
            row=6, column=2, columnspan=1, padx=15, pady=10, ipadx=10,
            ipady=13
        )

    def criar_estrutura(self):
        self.cria_pastas_de_trabalho()
        self.cria_bd()
        self.configura_bd()

    def altera_caminho(self, entrada, xlsx=False):
        if xlsx == True:
            caminho = filedialog.askopenfilename(
                initialdir=self.caminho_do_arquivo(),
                filetypes=(('Arquivos', '*.xlsx'), ("Tudo", '*.*'))
            )
        else:
            caminho = filedialog.askdirectory(
                initialdir=self.caminho_do_arquivo()
            )
        entrada.delete(0, 'end')
        entrada.insert(0, caminho)

    def atualizar_caminhos(self):
        resposta = messagebox.askyesno(
            ATUALIZAR_CAMINHOS[0], ATUALIZAR_CAMINHOS[1]
        )

        itens_para_atualizacao = [
            ['caminho_xlsx',
             '1',
             self.caminho_xlsx.get()],

            ['pasta_de_certidões',
             '2',
             self.caminho_pasta_de_certidões.get()],

            ['caminho_de_log',
             '3',
             self.caminho_log.get()],

            ['comprovantes_de_pagamento',
             '4',
             self.caminho_pasta_pagamento.get()],

            ['certidões_para_pagamento',
             '5',
             self.caminho_certidões_para_pagamento.get()]
        ]

        if resposta:
            arquivo = self.caminho_do_arquivo()
            with sqlite3.connect(f'{arquivo}/caminhos.db') as conexao:
                direcionador = conexao.cursor()
                for item in itens_para_atualizacao:
                    linha_update = (
                        f'UPDATE urls SET '
                        f'url = :{item[0]} WHERE oid = {item[1]}'
                    )
                    direcionador.execute(linha_update, {item[0]: item[2]})
                conexao.commit()
            self.janela_de_caminhos.destroy()
            messagebox.showinfo(
                CAMINHOS_ATUALIZADOS[0],
                (CAMINHOS_ATUALIZADOS[1])
            )
        else:
            self.janela_de_caminhos.destroy()

    def abrir_log(self):
        urls = self.consulta_urls()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        if not os.path.exists(f'{urls[2][1]}/{ano}-{mes}-{dia}.txt') \
                or (dia, mes, ano) == (' ', ' ', ' '):
            messagebox.showerror(LOG_INEXISTENTE[0], LOG_INEXISTENTE[1])
        else:
            caminho = f'{urls[2][1]}/{ano}-{mes}-{dia}.txt'
            novo_caminho = caminho.replace('/', '\\')
            os.startfile(novo_caminho)

    def cria_dias(self):
        contador_dia = 1
        while contador_dia <= 31:
            if contador_dia < 10:
                self.dias.append(f"0{contador_dia}")
                contador_dia += 1
            else:
                self.dias.append(str(contador_dia))
                contador_dia += 1

    def cria_meses(self):
        contador_mes = 1
        while contador_mes <= 12:
            if contador_mes < 10:
                self.meses.append(f"0{contador_mes}")
                contador_mes += 1
            else:
                self.meses.append(str(contador_mes))
                contador_mes += 1

    def cria_anos(self):
        contador_anos = 2010
        while contador_anos <= 2040:
            self.anos.append(str(contador_anos))
            contador_anos += 1

    def cria_calendario(self):
        self.cria_dias()
        self.cria_meses()
        self.cria_anos()

    #Continuar refatoramento a partir daqui.
    def checa_urls(self):
        urls = self.consulta_urls()
        if not os.path.exists(urls[0][1]):
            messagebox.showerror('Sumiu!!!', CHECA_URL_0)
        elif not os.path.exists(urls[1][1]):
            messagebox.showerror('Sumiu!!!', CHECA_URL_1)
        elif not os.path.exists(urls[2][1]):
            messagebox.showerror('Sumiu!!!', CHECA_URL_2)
        elif not os.path.exists(urls[3][1]):
            messagebox.showerror('Sumiu!!!', CHECA_URL_3)
        elif not os.path.exists(urls[4][1]):
            messagebox.showerror('Sumiu!!!', CHECA_URL_4)


    def executa(self):
        tempo_inicial = time.time()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if (not os.path.exists(urls[0][1])
                or not os.path.exists(urls[1][1])
                or not os.path.exists(urls[2][1])
                or not os.path.exists(urls[4][1])
                or not os.path.exists(urls[3][1])):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            objUniao = Uniao(dia, mes, ano)
            objTst = Tst(dia, mes, ano)
            objFgts = Fgts(dia, mes, ano)
            objGdf = Gdf(dia, mes, ano)
            lista_de_objetos = [objUniao, objTst, objFgts, objGdf]

            obj1.mensagem_de_log_completa(
                INICIO_DA_EXECUCAO, obj1.caminho_de_log
            )

            obj1.analisa_referencia()
            obj1.dados_completos_dos_fornecedores()
            obj1.listar_cnpjs()
            obj1.listar_cnpjs_exceções()

            obj1.mensagem_de_log_simples(ANALISADOS, obj1.caminho_de_log)

            for emp in obj1.empresas:
                obj1.mensagem_de_log_simples(f'{emp}', obj1.caminho_de_log)

            obj1.cria_diretorio()
            obj1.apaga_imagem()
            obj1.certidoes_n_encontradas()
            obj1.pdf_para_jpg()
            obj1.analisa_certidoes(lista_de_objetos)

            obj1.mensagem_de_log_simples(CONFERENCIA, obj1.caminho_de_log)

            obj1.pega_cnpj()

            obj1.mensagem_de_log_simples(PENDENCIAS, obj1.caminho_de_log)

            for emp in obj1.empresas_a_atualizar:
                obj1.mensagem_de_log_simples(
                    (f'{emp} - {obj1.empresas_a_atualizar[emp][0:-1]} '
                     f'- CNPJ: {obj1.empresas_a_atualizar[emp][-1]}\n'),
                    obj1.caminho_de_log
                )

            obj1.apaga_imagem()

            tempo_final = time.time()
            tempo_de_execução = int((tempo_final - tempo_inicial))

            obj1.mensagem_de_log_completa(
                (f'\n\nTempo total de execução: {tempo_de_execução // 60} '
                 f'minutos e {tempo_de_execução % 60} segundos.'),
                obj1.caminho_de_log
            )

            obj1.mensagem_de_log_simples(
                LINHA_FINAL, obj1.caminho_de_log)

            messagebox.showinfo(ANALISE_EXECUTADA[0], ANALISE_EXECUTADA[1])

    def selecionador_de_opções(self):
        print(self.variavel_de_opções.get())
        if self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[1]:
            self.pdf_para_jpg_para_renomear_arquivo()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[2]:
            self.pdf_para_jpg_renomear_conteudo_da_pasta()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[3]:
            self.renomeia()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[0]:
            messagebox.showwarning(
                OPCOES_DE_RENOMEACAO[4],
                OPCOES_DE_RENOMEACAO[5]
            )

    def renomeia(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if (
                not os.path.exists(urls[0][1])
                or not os.path.exists(urls[1][1])
                or not os.path.exists(urls[2][1])
                or not os.path.exists(urls[4][1])
                or not os.path.exists(urls[3][1])):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.apaga_imagem()
            obj1.pdf_para_jpg_renomear()
            obj1.gera_nome()
            obj1.apaga_imagem()

            messagebox.showinfo(
                RENOMEACAO_EXECUTADA[0],
                RENOMEACAO_EXECUTADA[1]
            )

    def transfere_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        urls = self.consulta_urls()
        if(
                not os.path.exists(urls[0][1])
                or not os.path.exists(urls[1][1])
                or not os.path.exists(urls[2][1])
                or not os.path.exists(urls[4][1])
                or not os.path.exists(urls[3][1])):
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
        if(
                not os.path.exists(urls[0][1])
                or not os.path.exists(urls[1][1])
                or not os.path.exists(urls[2][1])
                or not os.path.exists(urls[4][1])
                or not os.path.exists(urls[3][1])):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.merge()

    #continuar limpeza de constantes a partir daqui
    def pdf_para_jpg_para_renomear_arquivo(self):
        arquivo_selecionado = filedialog.askopenfilenames(
            initialdir=f'{self.caminho_do_arquivo()}/Certidões',
            filetypes=(('PDF', '*.pdf'), ("Tudo", '*.*'))
        )

        if list(arquivo_selecionado) == []:
            messagebox.showerror(
                ARQUIVOS_NAO_SELECIONADOS[0],
                ARQUIVOS_NAO_SELECIONADOS[1]
            )
            print(ARQUIVOS_NAO_SELECIONADOS[1])

        elif not os.path.exists(arquivo_selecionado[0]):
            print(ARQUIVO_INEXISTENTE[0])

            messagebox.showerror(
                ARQUIVO_INEXISTENTE[1],
                ARQUIVO_INEXISTENTE[2]
            )

        else:
            print(CRIANDO_IMAGENS[0])

            certidão_pdf = list(arquivo_selecionado)

            #lixo de debbug
            print(f'Arquivo que está sendo renomeado: {certidão_pdf}\n')

            for arquivo_a_renomear in certidão_pdf:
                ultima_barra: int = arquivo_a_renomear[::-1].find('/')+1
                os.chdir(arquivo_a_renomear[0:-(ultima_barra)])

                pages = convert_from_path(
                    arquivo_a_renomear, 300, last_page=1
                )

                imagem_da_certidao = f'{arquivo_a_renomear[:-4]}.jpg'
                pages[0].save(imagem_da_certidao, "JPEG")
                print(CRIANDO_IMAGENS[1])

                certidao_jpg = pytesseract.image_to_string(
                    Image.open(imagem_da_certidao), lang='por'
                )

                #lixo de debbug
                print('Renomeando certidão')

                for frase in IDENTIFICADOR_DE_CERTIDAO:
                    if frase in certidao_jpg:
                        if frase == 'GOVERNO DO DISTRITO FEDERAL':
                            try:
                                data = re.compile(IDENTIFICADOR_DE_VALIDADE_2[frase])
                                procura = data.search(certidao_jpg)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                            except AttributeError:
                                data = re.compile(IDENTIFICADOR_DE_VALIDADE[frase])
                                procura = data.search(certidao_jpg)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                        else:
                            data = re.compile(IDENTIFICADOR_DE_VALIDADE[frase])
                            procura = data.search(certidao_jpg)
                            datanome = procura.group()
                            separa = datanome.split('/')
                            junta = '-'.join(separa)
                        if ':' in junta:
                            retira = junta.split(':')
                            volta = ' '.join(retira)
                            junta = volta
                        shutil.move(
                            f'{imagem_da_certidao[0:-4]}.pdf',
                            f'{IDENTIFICADOR_TRADUZIDO[frase]} - {junta}.pdf'
                        )
                        os.unlink(imagem_da_certidao)
            print(
                '\nProcesso de renomeação de certidão executado com sucesso!'
            )
            print(
                '============================================================'
                '============================================================'
                '============================================================'
                '============================================================'
            )

            messagebox.showinfo(
                'Renomeou, miserávi!',
                'As certidões selecionadas foram renomeadas com sucesso!'
            )

    def caminho_de_pastas(self):
        pasta = 'Nenhuma pasta selecionada'
        self.pasta_selecionada = filedialog.askdirectory(
            initialdir=f'{self.caminho_do_arquivo()}/Certidões'
        )

        if (
                os.path.isdir(self.pasta_selecionada)
                and self.pasta_selecionada != (
                f'{self.caminho_do_arquivo()}/Certidões')
        ):

            pasta = self.pasta_selecionada
            self.caminho_da_pasta = Label(
                self.frame_renomear, text=os.path.basename(pasta), pady=0,
                padx=0, bg='white', fg='gray', font=('Helvetica', 9, 'bold')
            )

            self.caminho_da_pasta.grid(
                row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0,
                ipady=8, sticky=W+E
            )

        else:
            self.caminho_da_pasta = Label(
                self.frame_renomear, text=pasta, pady=0, padx=0, bg='white',
                fg='gray', font=('Helvetica', 9, 'bold')
            )

            self.caminho_da_pasta.grid(
                row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0,
                ipady=8, sticky=W+E
            )


    def apaga_imagens_da_pasta(self):
            os.chdir(self.pasta_selecionada)
            for arquivo in os.listdir(self.pasta_selecionada):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pasta_selecionada}/{arquivo}')

    def pdf_para_jpg_renomear_conteudo_da_pasta(self):
        self.pasta_selecionada = filedialog.askdirectory(
            initialdir=f'{self.caminho_do_arquivo()}/Certidões'
        )

        if (
                self.pasta_selecionada == 'Selecione a pasta que deseja renomear'
                or self.pasta_selecionada == ''
        ):
            messagebox.showerror(
                'Se não selecionar a pasta, não vai rolar!',
                ('Selecione uma pasta que contenha certidões que precisam ser'
                 ' renomeadas.')
            )
            print('nenhuma pasta selecionada')

            self.caminho_da_pasta = Label(
                self.frame_renomear, text='Nenhuma pasta selecionada', pady=0,
                padx=0, bg='white', fg='gray', font=('Helvetica', 9, 'bold')
            )

            self.caminho_da_pasta.grid(
                row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0,
                ipady=8, sticky=W+E
            )

        else:
            print(
                ('==========================================================='
                 '==========================================================='
                 '==========================================================='
                 '==========================================================='
                 )
            )

            print('\nCriando imagens:\n')
            os.chdir(self.pasta_selecionada)
            self.apaga_imagens_da_pasta()
            for pdf_file in os.listdir(self.pasta_selecionada):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(
                            f'{self.pasta_selecionada}/Mesclados'
                    ):
                        os.makedirs(f'{self.pasta_selecionada}/Mesclados')
                        shutil.move(
                            pdf_file,
                            f'{self.pasta_selecionada}/Mesclados/{pdf_file}'
                        )
                    else:
                        shutil.move(
                            pdf_file,
                            f'{self.pasta_selecionada}/Mesclados/{pdf_file}'
                        )

                elif pdf_file.endswith(".pdf"):
                    print(pdf_file[:-4])
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

            print(
                f'\nRenomeando certidões da pasta {self.pasta_selecionada}:\n'
                f'\n'
            )

            os.chdir(f'{self.pasta_selecionada}')
            origem = f'{self.pasta_selecionada}'
            for imagem in os.listdir(origem):
                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(
                        Image.open(f'{origem}/{imagem}'), lang='por'
                    )

                    for frase in IDENTIFICADOR_DE_CERTIDAO:
                        if frase in certidao:
                            if frase == 'GOVERNO DO DISTRITO FEDERAL':
                                try:
                                    data = re.compile(IDENTIFICADOR_DE_VALIDADE_2[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                                except AttributeError:
                                    data = re.compile(IDENTIFICADOR_DE_VALIDADE[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                            else:
                                data = re.compile(IDENTIFICADOR_DE_VALIDADE[frase])
                                procura = data.search(certidao)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                            if ':' in junta:
                                retira = junta.split(':')
                                volta = ' '.join(retira)
                                junta = volta
                            shutil.move(
                                f'{origem}/{imagem[0:-4]}.pdf',
                                f'{IDENTIFICADOR_TRADUZIDO[frase]} - {junta}.pdf'
                            )
                            print(imagem.split()[0])
            self.apaga_imagens_da_pasta()
            print(
                '\nProcesso de renomeação de certidões executado com sucesso!'
            )

            messagebox.showinfo(
                'Renomeou, miserávi!',
                ('Todas as certidões da pasta selecionada foram renomeadas co'
                 'm sucesso!'
                 )
            )

if __name__ == '__main__':
    tela = Tk()

    objeto_tela = Analisador(tela)
    tela.resizable(False, False)
    tela.title('GEOF - Analisador de certidões')
    #tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')
    tela.config(menu=objeto_tela.menu_certidões)

    tela.mainloop()
