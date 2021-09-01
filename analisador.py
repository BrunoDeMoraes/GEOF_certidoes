import os
import sqlite3
import threading
import time
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

from certidão import Certidao
from constantes import ANALISE_EXECUTADA
from constantes import ARQUIVOS_NAO_SELECIONADOS
from constantes import ATUALIZAR_CAMINHOS
from constantes import CAMINHOS_ATUALIZADOS
from constantes import CHECA_URL_0, CHECA_URL_1, CHECA_URL_2
from constantes import CHECA_URL_3, CHECA_URL_4
from constantes import LINHA_FINAL
from constantes import LOG_INEXISTENTE
from constantes import OPCOES_DE_RENOMEACAO, OPCOES_DE_RENOMEAR_VAZIA
from constantes import PASTA_NAO_SELECIONADA
from constantes import TEXTO_ANALISAR, TEXTO_CRIA_ESTRUTURA
from constantes import TEXTO_MESCLA_ARQUIVOS, TEXTO_PRINCIPAL
from constantes import TEXTO_RENOMEAR, TEXTO_TRANSFERE_ARQUIVOS
from constantes import TITULO_DA_INTERFACE
from fgts import Fgts
from gdf import Gdf
from tst import Tst
from união import Uniao


class Analisador(Certidao):
    def __init__(self, tela):
        if self.cria_bd():
            self.configura_bd()
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
            self.frame_data, text=TEXTO_PRINCIPAL, pady=3, padx=0, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Berlin Sans FB', 11)
        )

        self.dia_etiqueta = Label(
            self.frame_data, text='Dia', padx=20, pady=3, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Berlin Sans FB', 11)
        )
        self.mes_etiqueta = Label(
            self.frame_data, text='Mês', padx=20, pady=3, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Berlin Sans FB', 11)
        )
        self.ano_etiqueta = Label(
            self.frame_data, text='Ano', padx=20, pady=3, bg='green',
            fg='white', bd=2, relief=SUNKEN, font=('Berlin Sans FB', 11)
        )

        self.dia = StringVar()
        self.dia.set(" ")
        self.mes = StringVar()
        self.mes.set(" ")
        self.ano = StringVar()
        self.ano.set(" ")
        self.dias = [' ']
        self.meses = [' ']
        self.anos = [' ']
        self.cria_calendario()

        self.botao_abrir_log = Button(
            self.frame_data, text='Abrir log', command=self.abrir_log, padx=0,
            pady=0, bg='white', fg='green', font=('Helvetica', 9, 'bold'),
            bd=1
        )

        self.botao_abrir_log.bind("<Enter>", self.altera_botao_log)
        self.botao_abrir_log.bind("<Leave>", self.restaura_botao_log)

        self.validacao1 = OptionMenu(
            self.frame_data, self.dia, *self.dias)
        self.validacao2 = OptionMenu(
            self.frame_data, self.mes, *self.meses)
        self.validacao3 = OptionMenu(
            self.frame_data, self.ano, *self.anos)

        self.titulo_analisar = Label(
            self.frame_mestre, text=TEXTO_ANALISAR, pady=0, padx=0,
            bg='white', fg='black', font=('Arial Narrow', 11)
        )

        self.botao_analisar = Button(
            self.frame_mestre, text='Analisar\ncertidões',
            disabledforeground='white', command=self.thread_analisar,
            padx=30, pady=1, bg='green', fg='white',
            font=('Helvetica', 9, 'bold'), bd=1
        )
        self.botao_analisar.bind("<Enter>", self.altera_botao_analisar)
        self.botao_analisar.bind("<Leave>", self.restaura_botao_analisar)

        self.titulo_renomear = Label(
            self.frame_mestre, text=TEXTO_RENOMEAR, pady=0, padx=0,
            bg='white', fg='black', font=('Arial Narrow', 11)
        )

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set(OPCOES_DE_RENOMEACAO[0])
        self.validacao = OptionMenu(
            self.frame_mestre, self.variavel_de_opções, *OPCOES_DE_RENOMEACAO,
        )

        self.botao_renomear_tudo = Button(
            self.frame_mestre, text='Renomear\ncertidões',
            disabledforeground='white',
            command=self.thread_selecionador_de_opções, padx=30, pady=1,
            bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.botao_renomear_tudo.bind("<Enter>", self.altera_botao_renomear)
        self.botao_renomear_tudo.bind("<Leave>", self.restaura_botao_renomear)

        self.titulo_transfere_arquivos = Label(
            self.frame_mestre, text=TEXTO_TRANSFERE_ARQUIVOS,
            pady=0, padx=0, bg='white', fg='black',
            font=('Arial Narrow', 11)
        )

        self.botao_transfere_arquivos = Button(
            self.frame_mestre, text='Transferir\ncertidões',
            disabledforeground='white',
            command=self.thread_transfere_certidoes, padx=30, pady=1,
            bg='green', fg='white', font=('Helvetica', 9, 'bold'), bd=1
        )

        self.botao_transfere_arquivos.bind(
            "<Enter>", self.altera_botao_transfere
        )
        self.botao_transfere_arquivos.bind(
            "<Leave>", self.restaura_botao_transfere
        )

        self.titulo_mescla_arquivos = Label(
            self.frame_mestre, text=TEXTO_MESCLA_ARQUIVOS, pady=0, padx=0,
            bg='white', fg='black', font=('Arial Narrow', 11)
        )

        self.botao_mescla_arquivos = Button(
            self.frame_mestre, text='Mesclar\narquivos',
            command=self.thread_mescla_certidoes, padx=30, pady=1,
            disabledforeground='white', bg='green', fg='white',
            font=('Helvetica', 9, 'bold'), bd=1
        )

        self.botao_mescla_arquivos.bind(
            "<Enter>", self.altera_botao_mesclar
        )
        self.botao_mescla_arquivos.bind(
            "<Leave>", self.restaura_botao_mesclar
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

        self.validacao.grid(row=4, column=1, columnspan=7, ipadx=0, pady=10)

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

    def abrir_log(self):
        urls = self.consulta_urls()
        if (
                not os.path.exists(
                    f'{urls[2][1]}/{self.ano.get()}-{self.mes.get()}-'
                    f'{self.dia.get()}.txt'
                )
                or (
                (self.dia.get(), self.mes.get(), self.ano.get())
                == (' ', ' ', ' ')
        )
        ):
            messagebox.showerror(LOG_INEXISTENTE[0], LOG_INEXISTENTE[1])
        else:
            caminho = (
                f'{urls[2][1]}/{self.ano.get()}-{self.mes.get()}-'
                f'{self.dia.get()}.txt'
            )
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

    def altera_botao_log(self, evento):
        self.botao_abrir_log['fg'] = "#8FBC8F"

    def restaura_botao_log(self, evento):
        self.botao_abrir_log['fg'] = "green"

    def altera_botao_analisar(self, evento):
        self.botao_analisar['bg'] = "#8FBC8F"

    def restaura_botao_analisar(self, evento):
        self.botao_analisar['bg'] = "green"

    def altera_botao_renomear(self, evento):
        self.botao_renomear_tudo['bg'] = "#8FBC8F"

    def restaura_botao_renomear(self, evento):
        self.botao_renomear_tudo['bg'] = "green"

    def altera_botao_transfere(self, evento):
        self.botao_transfere_arquivos['bg'] = "#8FBC8F"

    def restaura_botao_transfere(self, evento):
        self.botao_transfere_arquivos['bg'] = "green"

    def altera_botao_mesclar(self, evento):
        self.botao_mescla_arquivos['bg'] = "#8FBC8F"

    def restaura_botao_mesclar(self, evento):
        self.botao_mescla_arquivos['bg'] = "green"

    def desabilita_botoes_de_execucao(self):
        self.menu_certidões.entryconfig('Configurações', state='disabled')
        self.botao_abrir_log['state'] = 'disabled'
        self.botao_analisar['state'] = 'disabled'
        self.botao_renomear_tudo['state'] = 'disabled'
        self.botao_transfere_arquivos['state'] = 'disabled'
        self.botao_mescla_arquivos['state'] = 'disabled'

    def habilita_botoes_de_execucao(self):
        self.menu_certidões.entryconfig('Configurações', state='normal')
        self.botao_abrir_log['state'] = 'normal'
        self.botao_analisar['state'] = 'normal'
        self.botao_renomear_tudo['state'] = 'normal'
        self.botao_transfere_arquivos['state'] = 'normal'
        self.botao_mescla_arquivos['state'] = 'normal'

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

    def cria_objetos(self):
        objUniao = Uniao(self.dia.get(), self.mes.get(), self.ano.get())
        objTst = Tst(self.dia.get(), self.mes.get(), self.ano.get())
        objFgts = Fgts(self.dia.get(), self.mes.get(), self.ano.get())
        objGdf = Gdf(self.dia.get(), self.mes.get(), self.ano.get())
        lista_de_objetos = [objUniao, objTst, objFgts, objGdf]
        return lista_de_objetos

    def cabecalho_de_execucao(self, corpo_da_funcao):
        urls = self.consulta_urls()
        if (
                not os.path.exists(urls[0][1])
                or not os.path.exists(urls[1][1])
                or not os.path.exists(urls[2][1])
                or not os.path.exists(urls[3][1])
                or not os.path.exists(urls[4][1])
        ):
            self.checa_urls()
        else:
            corpo_da_funcao()

    def corpo_de_execucao_de_analise(self):
        tempo_inicial = time.time()
        self.desabilita_botoes_de_execucao()
        try:
            lista_de_objetos = self.cria_objetos()
            analise = Certidao(self.dia.get(), self.mes.get(), self.ano.get())

            analise.analisa_lista_de_fornecedores()
            analise.analisa_requisitos()
            analise.analisa_certidoes(lista_de_objetos)
            analise.resultado_da_analise()

            messagebox.showinfo(ANALISE_EXECUTADA[0], ANALISE_EXECUTADA[1])

            tempo_final = time.time()
            tempo_de_execução = int((tempo_final - tempo_inicial))

            analise.mensagem_de_log_completa(
                (f'\n\nTempo total de execução: {tempo_de_execução // 60} '
                 f'minutos e {tempo_de_execução % 60} segundos.'),
                analise.caminho_de_log
            )

            analise.mensagem_de_log_simples(
                LINHA_FINAL, analise.caminho_de_log)

        finally:
            self.habilita_botoes_de_execucao()

    def analisar_certidoes(self):
        self.cabecalho_de_execucao(self.corpo_de_execucao_de_analise)

    def thread_analisar(self):
        threading.Thread(target=self.analisar_certidoes).start()

    def renomear_certidoes_arquivos(self):
        self.desabilita_botoes_de_execucao()
        try:
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

            else:
                self.pdf_para_jpg_para_renomear_arquivos(arquivo_selecionado)
        finally:
            self.habilita_botoes_de_execucao()

    def renomear_certidoes_de_uma_pasta(self):
        self.desabilita_botoes_de_execucao()
        try:
            self.pasta_selecionada = filedialog.askdirectory(
                initialdir=f'{self.caminho_do_arquivo()}/Certidões'
            )

            if (
                    self.pasta_selecionada == (
                    'Selecione a pasta que deseja renomear'
            )
                    or self.pasta_selecionada == ''
            ):
                messagebox.showerror(
                    PASTA_NAO_SELECIONADA[0], PASTA_NAO_SELECIONADA[1]
                )

                print(PASTA_NAO_SELECIONADA[1])

            else:
                self.pdf_para_jpg_renomear_conteudo_da_pasta()
        finally:
            self.habilita_botoes_de_execucao()

    def corpo_de_execucao_de_renomeacao_geral(self):
        self.desabilita_botoes_de_execucao()
        try:
            renomeacao_geral = Certidao(
                self.dia.get(), self.mes.get(), self.ano.get()
            )
            renomeacao_geral.rotina_de_renomeacao_geral()
        finally:
            self.habilita_botoes_de_execucao()

    def renomear_certidoes_geral(self):
        self.cabecalho_de_execucao(self.corpo_de_execucao_de_renomeacao_geral)

    def selecionador_de_opções(self):
        if self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[1]:
            self.renomear_certidoes_arquivos()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[2]:
            self.renomear_certidoes_de_uma_pasta()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[3]:
            self.renomear_certidoes_geral()
        elif self.variavel_de_opções.get() == OPCOES_DE_RENOMEACAO[0]:
            messagebox.showwarning(
                OPCOES_DE_RENOMEAR_VAZIA[0],
                OPCOES_DE_RENOMEAR_VAZIA[1]
            )

    def thread_selecionador_de_opções(self):
        threading.Thread(target=self.selecionador_de_opções).start()

    def transfere_certidoes(self):
        self.desabilita_botoes_de_execucao()
        try:
            transferencia = Certidao(
                self.dia.get(), self.mes.get(), self.ano.get()
            )
            transferencia.rotina_de_transferencia_de_certidoes()
        finally:
            self.habilita_botoes_de_execucao()

    def thread_transfere_certidoes(self):
        threading.Thread(target=self.transfere_certidoes).start()

    def mescla_certidoes(self):
        self.desabilita_botoes_de_execucao()
        try:
            mescla = Certidao(
                self.dia.get(), self.mes.get(), self.ano.get()
            )
            mescla.rotina_de_mesclagem()
        finally:
            self.habilita_botoes_de_execucao()

    def thread_mescla_certidoes(self):
        threading.Thread(target=self.mescla_certidoes).start()

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
            bg='white', fg='green', font=('Arial Narrow', 11)
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
                self.altera_caminho(self.caminho_certidões_para_pagamento)
            ),
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

if __name__ == '__main__':
    tela = Tk()

    objeto_tela = Analisador(tela)
    tela.resizable(False, False)
    tela.title(TITULO_DA_INTERFACE)
    tela.config(menu=objeto_tela.menu_certidões)

    tela.mainloop()
