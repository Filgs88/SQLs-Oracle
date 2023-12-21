import xlsxwriter
from ConexaoOracle import Conexao

def sql_ordem_servico():
    sql = """
    with tabela1 as(
    select item_ordem_servico_tarefa.NUMERO_ORDEM_SERVICO
    ,coalesce(sum(item_ordem_servico_tarefa.horas),0) as horas_tarefa

    from industria.item_ordem_servico_tarefa

    where item_ordem_servico_tarefa.ano_ordem_servico >= 2023

    group by item_ordem_servico_tarefa.NUMERO_ORDEM_SERVICO)

    , tabela2 as(
    select item_ordem_servico_apont.NUMERO_ORDEM_SERVICO
    ,coalesce(sum(item_ordem_servico_apont.TOTHORAS_TRABALHADAS),0) as horas_apont

    from industria.item_ordem_servico_apont

    where item_ordem_servico_apont.ano_ordem_servico >= 2023

    group by item_ordem_servico_apont.NUMERO_ORDEM_SERVICO)

    select ordem_servico.ano_ordem_servico
        ,ordem_servico.NUMERO_ORDEM_SERVICO
        ,ordem_servico.CODIGO_CTE_ALFA
        ,ordem_servico.cod_funcionario
        ,coalesce(tabela1.horas_tarefa,0) as horas_tarefa
        ,coalesce(tabela2.horas_apont,0) as horas_apont
        ,ordem_servico.CODIGO_DESTINO_MANUTENCAO
        ,ordem_servico.CODIGO_TIPOMANUTENCAO
        ,ordem_servico.codigo_prioridade
        ,ordem_servico.SOLITACAO_SERVICO
        ,ordem_servico.data_abertura
        ,ordem_servico.DATA_ENCERRAMENTO
        ,ordem_servico.DATAHORA_ACEITE

    from industria.ordem_servico
    ,tabela1
    ,tabela2

    where ordem_servico.ano_ordem_servico >= 2023
    and tabela1.NUMERO_ORDEM_SERVICO (+)= ordem_servico.NUMERO_ORDEM_SERVICO
    and tabela2.NUMERO_ORDEM_SERVICO (+)= ordem_servico.NUMERO_ORDEM_SERVICO

    
    order by ordem_servico.numero_ordem_servico
    """

    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('OrdemServico.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'ANO_ORDEM_SERVICO')
    ws.write(r, c+1, 'NUMERO_ORDEM_SERVICO')
    ws.write(r, c+2, 'CODIGO_CTE_ALFA')
    ws.write(r, c+3, 'COD_FUNCIONARIO')
    ws.write(r, c+4, 'HORAS_TAREFA')
    ws.write(r, c+5, 'HORAS_APONT')
    ws.write(r, c+6, 'CODIGO_DESTINO_MANUTENCAO')
    ws.write(r, c+7, 'CODIGO_TIPOMANUTENCAO')
    ws.write(r, c+8, 'CODIGO_PRIORIDADE')
    ws.write(r, c+9, 'SOLITACAO_SERVICO')
    ws.write(r, c+10, 'DATA_ABERTURA')
    ws.write(r, c+11, 'DATA_ENCERRAMENTO')
    ws.write(r, c+12, 'DATAHORA_ACEITE')

    for col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13 in linhas:
        r += 1
        ws.write(r, c, col1)
        ws.write(r, c+1, col2)
        ws.write(r, c+2, col3)
        ws.write(r, c+3, col4)
        ws.write(r, c+4, col5)
        ws.write(r, c+5, col6)
        ws.write(r, c+6, col7)
        ws.write(r, c+7, col8)
        ws.write(r, c+8, col9)
        ws.write(r, c+9, col10)
        ws.write(r, c+10, col11)
        ws.write(r, c+11, col12)
        ws.write(r, c+12, col13)
        
    wb.close()
    cursor.close()
    con.close()

    print("Database Ordem de Serviço Atualizada!")
