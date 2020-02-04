from certidao2 import Certidao, Uniao, Tst, Fgts, Gdf
import time

tempo_inicial = time.time()
dia = '31'
mes = '01'
ano = '2020'
orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
empresasdic = {}

obj1 = Certidao(dia, mes, ano)
obj1.mensagem_log('Início da execução')
lis_ref_cel = obj1.pega_referencia()
if len(lis_ref_cel) == 0:
    obj1.mensagem_log('\nData específicada não encontrada')
    raise Exception('Data não encontrada!')
elif len(lis_ref_cel) > 1:
    obj1.mensagem_log('Data informada em multiplicidade')
    raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
          f'\nApague os valores duplicados e execute o programa novamente.')
else:
    ref_cel = lis_ref_cel[0]
    obj1.mensagem_log(f'\nReferência encontrada na célula {lis_ref_cel[0]}')
fornecedores = obj1.pega_fornecedores(ref_cel)
obj1.mensagem_log_sem_horario('\nFornecedores analisados:')
for emp in fornecedores:
    obj1.mensagem_log_sem_horario(f'{emp}')
obj1.cria_diretorio(fornecedores)
obj1.certidoes_n_encontradas(fornecedores, orgaos)
obj1.pdf_para_jpg(fornecedores, orgaos)
objUniao = Uniao(dia, mes, ano)
objTst = Tst(dia, mes, ano)
objFgts = Fgts(dia, mes, ano)
objGdf = Gdf(dia, mes, ano)
lista_objetos = [objUniao, objTst, objFgts, objGdf]

obj1.mensagem_log('\nInicio da conferência de datas de emissão e vencimento:')
for emp in fornecedores:
    empresadic = {}
    index = 0
    for objeto in lista_objetos:
        cert = objeto.pega_string(emp)
        print(f'Processando string {emp}')
        obj1.mensagem_log(f'\n{emp}')
        val = objeto.confere_data(cert)
        if val == True:
            empresadic[orgaos[index]] = 'OK'
        else:
            empresadic[orgaos[index]] = 'INCOMPATÍVEL'
        index += 1
    empresasdic[emp] = empresadic
print('RESULTADO DA CONFERÊNCIA:')
obj1.mensagem_log_sem_horario('\nRESULTADO DA CONFERÊNCIA:')
numerador = 0
for emp in empresasdic:
    print(f'{numerador + 1 :>2} - {emp}')
    print(f'{empresasdic[emp]}')
    obj1.mensagem_log_sem_horario(f'{numerador + 1 :>2} - {emp :^30} - {empresasdic[emp]}\n')
    numerador += 1
obj1.apaga_imagem(fornecedores)
tempo_final = time.time()
tempo_de_execução = int((tempo_final - tempo_inicial))
obj1.mensagem_log(f'\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
obj1.mensagem_log_sem_horario('\n\n======================================================================================\n')
