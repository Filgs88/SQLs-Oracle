import xlsxwriter
from ConexaoOracle import Conexao

def sql_solicitacao_compra():
    sql = """
    select solicitacaocompra.NR_SOLICITACAO
    ,COTACAOXSOLICITACAO.NR_COTACAO
    ,solicitacaocompra.COD_FUNCIONARIO
    ,solicitacaocompra.DATA
    ,coalesce(solicitacaocompra.QTDESOLICITADA,0) as QTDESOLICITADA
    ,coalesce(solicitacaocompra.QTDE_PENDENTE,0) as QTDE_PENDENTE
    ,coalesce(solicitacaocompra.QTDE_APROVACAO,0) as QTDE_APROVACAO
    ,solicitacaocompra.COD_MATERIAL
    ,material.descricao
    ,solicitacaocompra.OBSERVACAO
    ,solicitacaocompra.SOLICITACAOAPROVADA
    ,solicitacaocompra.SOLICITACAOPENDENTE
    ,solicitacaocompra.SITUACAO
    ,solicitacaocompra.INFORMACAO_FORNECEDOR
    ,solicitacaocompra.DATACRIACAO
    ,APROVACAOPARACOMPRA.dataaprovacao as aprovacaocompra
    ,APROVACAOSOLICITACAOCOMPRA.dataaprovacao as aprovacaosolicitacao
    ,solicitacaocompra.COD_ALMOXARIFADO

    from MATERIAL.solicitacaocompra
    ,MATERIAL.APROVACAOPARACOMPRA
    ,MATERIAL.APROVACAOSOLICITACAOCOMPRA
    ,material.material
    ,material.COTACAOXSOLICITACAO

    where solicitacaocompra.DATA > to_date('01/02/2023','dd/mm/yyyy')
    and solicitacaocompra.COD_ALMOXARIFADO in ('8','27')
    and APROVACAOPARACOMPRA.NR_SOLICITACAO (+)= solicitacaocompra.NR_SOLICITACAO
    and APROVACAOSOLICITACAOCOMPRA.NR_SOLICITACAO (+)= solicitacaocompra.NR_SOLICITACAO
    and solicitacaocompra.cod_material (+)= material.cod_material
    and COTACAOXSOLICITACAO.NR_SOLICITACAO (+)= solicitacaocompra.NR_SOLICITACAO
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('//192.168.1.177/Users/PC/OneDrive - MSFT/PCM/01. PCMI/25. SQLs/SolicitacaoCompra.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'NR_SOLICITACAO')
    ws.write(r, c+1, 'NR_COTACAO')
    ws.write(r, c+2, 'COD_FUNCIONARIO')
    ws.write(r, c+3, 'DATA')
    ws.write(r, c+4, 'QTDESOLICITADA')
    ws.write(r, c+5, 'QTDE_PENDENTE')
    ws.write(r, c+6, 'QTDE_APROVACAO')
    ws.write(r, c+7, 'COD_MATERIAL')
    ws.write(r, c+8, 'DESCRICAO')
    ws.write(r, c+9, 'OBSERVACAO')
    ws.write(r, c+10, 'SOLICITACAOAPROVADA')
    ws.write(r, c+11, 'SOLICITACAOPENDENTE')
    ws.write(r, c+12, 'SITUACAO')
    ws.write(r, c+13, 'INFORMACAO_FORNECEDOR')
    ws.write(r, c+14, 'DATACRIACAO')
    ws.write(r, c+15, 'APROVACAOCOMPRA')
    ws.write(r, c+16, 'APROVACAOSOLICITACAO')
    ws.write(r, c+17, 'COD_ALMOXARIFADO')

    for col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13, col14, col15, col16, col17, col18 in linhas:
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
        ws.write(r, c+13, col14)
        ws.write(r, c+14, col15)
        ws.write(r, c+15, col16)
        ws.write(r, c+16, col17)
        ws.write(r, c+17, col18)

    wb.close()
    cursor.close()
    con.close()

    print("Database Solicitação de Compra Atualizada!")

