from diretorios import Diretorio

pasta_mae = '//hrg-74977/GEOF/CERTIDÕES/CERTIDÕES GERAL - Cópia'

checagem = Diretorio(pasta_mae)
lista_de_pastas = checagem.lista_pastas()
print(lista_de_pastas)
lista_de_arqui = checagem.lista_arquivos(lista_de_pastas)
print(lista_de_arqui)
checagem.renomeia_arquivos(lista_de_arqui, ('data teste'))