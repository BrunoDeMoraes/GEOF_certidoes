from tkinter import *
from tkinter import filedialog
from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time

class Analisador:
    def __init__(self, tela):
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)

        #self.menu_certidões = Menu(tela)
        #self.tela.config(menu=menu_certidões)
        #self.menu_configurações = Menu(menu_certidões)
        #self.menu_certidões.add_cascade(label='Configurações', menu=menu_configurações)
        #self.menu_configurações.add_separator()
        #self.menu_configurações.add_command(label='Caminhos', command=abrir_janela_caminhos)
        #self.menu_configurações.add_separator()

        self.titulo = Label(self.frame_mestre, text='''Indique a data limite pretendida para o próximo pagamento e
        em seguida escolha uma das seguintes opções:''', pady=0, padx=0, bg='green', fg='white',
                       font=('Helvetica', 12, 'bold'))

        self.dia_etiqueta = Label(self.frame_mestre, text='Dia', padx=8, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 9, 'bold'))
        self.mes_etiqueta = Label(self.frame_mestre, text='Mês', padx=8, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 9, 'bold'))
        self.ano_etiqueta = Label(self.frame_mestre, text='Ano', padx=8, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
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

        self.validacao = OptionMenu(self.frame_mestre, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(self.frame_mestre, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(self.frame_mestre, self.variavel3, *self.anos)

        self.titulo_analisar = Label(self.frame_mestre, text='''Utilize esta opção para identificar quais certidões devem ser atualizadas
        ou se há requisitos a cumprir para a devida execução da análise.''', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_analisar = Button(self.frame_mestre, text='Analisar\ncertidões', command=self.executa, padx=25, pady=10, bg='green',
                                fg='white', font=('Helvetica', 11, 'bold'), bd=1)

        self.titulo_renomear = Label(self.frame_mestre, text='''Após atualizar as certidões, use esta opção para padronizar os nomes dos 
        arquivos e em seguida faça nova análise para certificar que está tudo OK.''', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_renomear = Button(self.frame_mestre, text='Renomear\ncertidões', command=self.renomeia, padx=20, pady=10,
                                bg='green',
                                fg='white', font=('Helvetica', 11, 'bold'), bd=1)

        self.titulo_transfere_arquivos = Label(self.frame_mestre, text='''Esta opção transfere as certidões para uma pasta identificada pela data
        do pagamento. Esse passo deve ser executado logo após a análise.''', pady=0, padx=0, bg='white', fg='black',
                                          font=('Helvetica', 9, 'bold'))

        self.botao_transfere_arquivos = Button(self.frame_mestre, text='Transferir\ncertidões', command=self.transfere_certidoes,
                                          padx=20, pady=10, bg='green',
                                          fg='white', font=('Helvetica', 11, 'bold'), bd=1)

        self.titulo_mescla_arquivos = Label(self.frame_mestre, text='''Se já houve o pagamento e os comprovantes estão na devida pasta, esta 
        opção mescla os comprovantes com suas respectivas certidões.''', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))

        self.botao_mescla_arquivos = Button(self.frame_mestre, text=' Mesclar  \narquivos', command=self.mescla_certidoes, padx=20,
                                       pady=10, bg='green',
                                       fg='white', font=('Helvetica', 11, 'bold'), bd=1)

        self.roda_pe = Label(self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
                        font=('Helvetica', 8, 'italic'), anchor=E)

        self.titulo.grid(row=0, column=1, columnspan=5, pady=10, sticky=W + E)
        self.dia_etiqueta.grid(row=1, column=1, pady=0, ipadx=14, ipady=0)
        self.mes_etiqueta.grid(row=1, column=2, pady=0, ipadx=14, ipady=0)
        self.ano_etiqueta.grid(row=1, column=3, pady=0, ipadx=14, ipady=0)
        self.validacao.grid(row=2, column=1, pady=0)
        self.validacao2.grid(row=2, column=2, pady=0)
        self.validacao3.grid(row=2, column=3, pady=0)
        self.titulo_analisar.grid(row=3, column=4, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_analisar.grid(row=3, column=5, padx=0, pady=10)
        self.titulo_renomear.grid(row=4, column=4, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_renomear.grid(row=4, column=5, padx=0, pady=10)
        self.titulo_transfere_arquivos.grid(row=5, column=4, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13,
                                       sticky=W + E)
        self.botao_transfere_arquivos.grid(row=5, column=5, padx=0, pady=10)
        self.titulo_mescla_arquivos.grid(row=6, column=4, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.botao_mescla_arquivos.grid(row=6, column=5, padx=0, pady=10)
        self.roda_pe.grid(row=7, column=1, columnspan=10, pady=5, sticky=W + E)


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

    def caminho_de_pastas():
        #pasta = filedialog.askdirectory(initialdir='C:/Users/14343258/Desktop')
        pass


    def abrir_janela_caminhos():
        janela_de_caminhos = Toplevel()
        janela_de_caminhos.title('Lista de caminhos')
        janela_de_caminhos.geometry('400x400')
        frame = LabelFrame(janela_de_caminhos, padx=0, pady=0)
        frame.pack(padx=1, pady=1)
        botao_caminho = Button(frame, text='Buscar caminho', command=caminho_de_pastas, padx=10,
                                       pady=10, bg='green',
                                       fg='white', font=('Helvetica', 11, 'bold'), bd=1)
        botao_caminho.grid(row=1, column=1, columnspan=1, padx = 15, pady=10, ipadx=10, ipady=13, sticky=W+E)
        titulo = Label(frame, text='Nada aqui', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))
        titulo.grid(row=1, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)

tela = Tk()

objeto_tela = Analisador(tela)
tela.resizable(False, False)
tela.title('GEOF - Analisador de certidões')
#tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')


tela.mainloop()
