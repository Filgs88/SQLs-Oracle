import xlsxwriter
from ConexaoOracle import Conexao

def sql_ordem_compra():
    sql = """
    Select itensordemcompra.nr_solicitacao
    ,ordemcompra.dataoc
    ,ordemcompra.nroc
    ,ordemcompra.nr_cotacao
    ,ITENSENTRADA.sequencia_nf
    ,ITENSENTRADA.nrnf
    ,ordemcompra.cod_fornecedor
    ,ordemcompra.cod_plano
    ,itensordemcompra.cod_material
    ,solicitacaocompra.observacao
    ,solicitacaocompra.cod_almoxarifado

    from material.ORDEMCOMPRA
    ,material.ITENSORDEMCOMPRA
    ,material.solicitacaocompra
    ,material.ITENSENTRADA

    where ordemcompra.dataoc > to_date('01/03/2023','dd/mm/yyyy')
    and solicitacaocompra.cod_almoxarifado in ('8', '27')
    and ordemcompra.nroc = itensordemcompra.nroc
    and itensordemcompra.nr_solicitacao = solicitacaocompra.nr_solicitacao
    and itensentrada.nroc (+)= ordemcompra.nroc
    
    order by ordemcompra.dataoc
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('//192.168.1.177/Users/PC/OneDrive - MSFT/PCM/01. PCMI/25. SQLs/OrdemCompra.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'NR_SOLICITACAO')
    ws.write(r, c+1, 'DATAOC')
    ws.write(r, c+2, 'NROC')
    ws.write(r, c+3, 'NR_COTACAO')
    ws.write(r, c+4, 'SEQUENCIA_NF')
    ws.write(r, c+5, 'NRNF')
    ws.write(r, c+6, 'COD_FORNECEDOR')
    ws.write(r, c+7, 'COD_PLANO')
    ws.write(r, c+8, 'COD_MATERIAL')
    ws.write(r, c+9, 'OBSERVACAO')
    ws.write(r, c+10, 'COD_ALMOXARIFADO')

    for col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11 in linhas:
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
        
    wb.close()
    cursor.close()
    con.close()

    print("Database Ordem de Compra Atualizada!")