import os
import sqlite3

class Conexao:
    def caminho_do_arquivo(self):
        caminho_py = __file__
        caminho_do_dir = caminho_py.split('\\')
        caminho_de_uso = ('/').join(caminho_do_dir[0:-1])
        return caminho_de_uso

    def caminho_do_bd(self):
        caminho = self.caminho_do_arquivo()
        banco_de_dados = f'{caminho}/caminhos.db'
        return banco_de_dados

    def cria_pastas_de_trabalho(self):
        caminho = self.caminho_do_arquivo()
        pastas_de_trabalho = [
            'Certidões',
            'Logs de conferência',
            'Certidões para pagamento',
            'Comprovantes de pagamento'
        ]
        for pasta in pastas_de_trabalho:
            if not os.path.exists(f'{caminho}/{pasta}'):
                os.makedirs(f'{caminho}/{pasta}')
                print(f'Pasta de trabalho {pasta} criada com sucesso!\n')
            else:
                print(f'Pasta de trabalho {pasta} localizada.\n')

    def conexao(self, comando):
        caminho_do_banco_de_dados = self.caminho_do_bd()
        with sqlite3.connect(caminho_do_banco_de_dados) as conexao:
            direcionador = conexao.cursor()
            direcionador.execute(comando)

    def consulta_urls(self):
        caminho_do_banco_de_dados = self.caminho_do_bd()
        print(
            (
                f'caminho de banco de dados dentro da função consulta '
                f'{caminho_do_banco_de_dados}'
            )
        )
        with sqlite3.connect(caminho_do_banco_de_dados) as conexao:
            comando = 'SELECT *, oid FROM urls'
            direcionador = conexao.cursor()
            direcionador.execute(comando)
            urls = direcionador.fetchall()
            for registro in urls:
                print(f'{registro[0]}: {registro[1]}\n')
            return urls

    def cria_bd(self):
        caminho_do_banco_de_dados = self.caminho_do_bd()
        if not os.path.exists(caminho_do_banco_de_dados):
            comando = 'CREATE TABLE urls (variavel text, url text)'
            self.conexao(comando)
        else:
            print('Banco de dados localizado.')
            self.consulta_urls()

    def configura_bd(self):
        caminho = self.caminho_do_arquivo()
        enderecos = {'caminho_xlsx': f'{caminho}/listas.xlsx',
                    'pasta_de_certidões': f'{caminho}/Certidões',
                    'caminho_de_log': f'{caminho}/Logs de conferência',
                    'comprovantes_de_pagamento': (
                        f'{caminho}/Comprovantes de pagamento'
                    ),
                    'certidões_para_pagamento': (f'{caminho}/Certidões para pagamento')
                     }
        for endereco in enderecos:
            comando = 'INSERT INTO urls VALUES (:variavel, :url)'
            substituto = {"variavel": endereco, "url": enderecos[endereco]}
            caminho_do_banco_de_dados = self.caminho_do_bd()
            with sqlite3.connect(caminho_do_banco_de_dados) as conexao:
                direcionador = conexao.cursor()
                direcionador.execute(comando, substituto)
        self.consulta_urls()
