import xlsxwriter
from ConexaoOracle import Conexao

def sql_solicitacao_servico():
    sql = """
    SELECT ordem_servico.ano_ordem_servico
    ,ordem_servico.numero_ordem_servico
    ,ordem_servico.data_abertura
    ,ordem_servico.solitacao_servico
    ,requisicaomaterial.nrrequisicao
    ,itensrequisicaomaterial.nr_solicitacao
    ,ordem_servico.codigo_destino_manutencao
    ,material.cod_material
    ,material.descricao
    ,coalesce(itensrequisicaomaterial.qtdesolicitada,0) as qtdesolicitada
    ,itensrequisicaomaterial.observacao
    ,ordem_servico.cod_objetocusto
    ,ordem_servico.codigo_cte_alfa

    FROM industria.ordem_servico
    ,material.itensrequisicaomaterial
    ,material.requisicaomaterial
    ,material.material

    WHERE ordem_servico.numero_ordem_servico (+)= requisicaomaterial.numero_ordem_servico
    and itensrequisicaomaterial.cod_material = material.cod_material
    and itensrequisicaomaterial.nrrequisicao = requisicaomaterial.nrrequisicao
    and ordem_servico.ano_ordem_servico in ('2023')
    and ordem_servico.codigo_destino_manutencao in ('12', '109')
    and material.cod_unidade in ('SV')
    and requisicaomaterial.datarequisicao > to_date('01/01/2023', 'dd/mm/yyyy')

    ORDER BY ordem_servico.numero_ordem_servico
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('//192.168.1.177/Users/PC/OneDrive - MSFT/PCM/01. PCMI/25. SQLs/SolicitacaoServico.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'ANO_ORDEM_SERVICO')
    ws.write(r, c+1, 'NUMERO_ORDEM_SERVICO')
    ws.write(r, c+2, 'DATA_ABERTURA')
    ws.write(r, c+3, 'SOLICITACAO_SERVICO')
    ws.write(r, c+4, 'NRREQUISICAO')
    ws.write(r, c+5, 'NR_SOLICITACAO')
    ws.write(r, c+6, 'CODIGO_DESTINO_MANUTENCAO')
    ws.write(r, c+7, 'COD_MATERIAL')
    ws.write(r, c+8, 'DESCRICAO')
    ws.write(r, c+9, 'QTDESOLICITADA')
    ws.write(r, c+10, 'OBSERVACAO')
    ws.write(r, c+11, 'COD_OBJETOCUSTO')
    ws.write(r, c+12, 'CODIGO_CTE_ALFA')

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

    print("Database Solcitação de Serviços Atualizada!")

