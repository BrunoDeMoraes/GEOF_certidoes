import threading
from tkinter import *
from tkinter import ttk


class Barra:
    def barra_de_progresso(self, titulo, base):
        self.janela_da_barra_de_progresso = Toplevel()
        self.janela_da_barra_de_progresso.title(titulo)
        self.janela_da_barra_de_progresso.resizable(0, 0)
        self.cria_barra(base)

    def cria_barra(self, base):
        self.barra = ttk.Progressbar(
            self.janela_da_barra_de_progresso, orient=HORIZONTAL, length=400,
            mode='determinate'
        )
        self.barra.pack(padx=20, pady=10)

        self.valor_executado = Label(
            self.janela_da_barra_de_progresso,
            text=f'Total executado {base:.2f}%'
        )
        self.valor_executado.pack()

    def valor_da_barra(self, base):
        self.barra['value'] = base
        self.valor_executado.destroy()
        self.valor_executado = Label(
            self.janela_da_barra_de_progresso,
            text=f'Total executado {base:.2f}%'
        )
        self.valor_executado.pack()

    def thread_barra_de_progresso(self, titulo, base):
        threading.Thread(
            target=lambda: self.barra_de_progresso(titulo, base)
        ).start()

    def destruir_barra_de_progresso(self):
        self.janela_da_barra_de_progresso.destroy()
