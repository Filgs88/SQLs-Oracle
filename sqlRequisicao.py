import xlsxwriter
from ConexaoOracle import Conexao

def sql_requisicao():
    sql = """
    select NRREQUISICAO
    ,ITEM
    ,COD_MATERIAL
    ,coalesce(quantidade,0) as quantidade
    ,DATARETIRADA
    ,coalesce(VRCUSTOUNITARIO,0) as VRCUSTOUNITARIO
    ,coalesce(QTDESOLICITADA,0) as QTDESOLICITADA
    ,CODIGO_CTE
    ,COD_FUNCIONARIO_RETIRAR
    ,COD_ITEM_CUSTO
    ,COD_OBJETOCUSTO_ATIVO
    ,JUSTIFICATIVA_CANCELAM

    from material.ITENSREQUISICAOMATERIAL

    where datarequisicao_item > to_date('01/01/2023','dd/mm/yyyy')
    and cod_almoxarifado = 8
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('Requisição.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'NRREQUISICAO')
    ws.write(r, c+1, 'ITEM')
    ws.write(r, c+2, 'COD_MATERIAL')
    ws.write(r, c+3, 'QUANTIDADE')
    ws.write(r, c+4, 'DATARETIRADA')
    ws.write(r, c+5, 'VRCUSTOUNITARIO')
    ws.write(r, c+6, 'QTDESOLICITADA')
    ws.write(r, c+7, 'CODIGO_CTE')
    ws.write(r, c+8, 'COD_FUNCIONARIO_RETIRAR')
    ws.write(r, c+9, 'COD_ITEM_CUSTO')
    ws.write(r, c+10, 'COD_OBJETOCUSTO_ATIVO')
    ws.write(r, c+11, 'JUSTIFICATIVA_CANCELAM')

    for col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 in linhas:
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

    wb.close()
    cursor.close()

    con.close()

    print("Database Requisições Atualizada!")

