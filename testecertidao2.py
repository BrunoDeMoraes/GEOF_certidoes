from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf

dia = '02'
mes = '09'
ano = '2019'
orgaos = ['UNIAO', 'TST', 'FGTS', 'GDF']
empresasdic = {}


obj1 = Certidao(dia, mes, ano)
lis_ref_cel = obj1.pega_referencia()
if len(lis_ref_cel) == 0:
    raise Exception('Data não encontrada!')
elif len(lis_ref_cel) > 1:
    raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
          f'\nApague os valores duplicados e execute o programa novamente.')
else:
    ref_cel = lis_ref_cel[0]
    obj1.mensagem_log(f'Referência encontrada na céclula {lis_ref_cel[0]}')
fornecedores = obj1.pega_fornecedores(ref_cel)
obj1.mensagem_log('\n\nFornecedores analisados:')
for emp in fornecedores:
    obj1.mensagem_log(f'\n{emp}')
obj1.cria_diretorio(fornecedores)
obj1.certidoes_n_encontradas(fornecedores, orgaos)
obj1.pdf_para_jpg(fornecedores, orgaos)
objUniao = Uniao(dia, mes, ano)
objTst = Tst(dia, mes, ano)
objFgts = Fgts(dia, mes, ano)
objGdf = Gdf(dia, mes, ano)
lista_objetos = [objUniao, objTst, objFgts, objGdf]

for emp in fornecedores:
    empresadic = {}
    index = 0
    for objeto in lista_objetos:
        cert = objeto.pega_string(emp)
        print(f'Tudo certo até aqui {emp} {objeto}')
        objeto.mensagem_log(f'\npra {emp}')
        val = objeto.confere_data(cert)
        if val == True:
            empresadic[orgaos[index]] = 'OK'
        else:
            empresadic[orgaos[index]] = 'Certidão não compreende data de pagamento'
        index += 1
    empresasdic[emp] = empresadic
print('Resultado da conferência:')
numerador = 0
for emp in empresasdic:
    print(f'{numerador} - {emp} - {empresasdic[emp]}')
    numerador += 1