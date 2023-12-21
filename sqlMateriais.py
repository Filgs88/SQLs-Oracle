import xlsxwriter
from ConexaoOracle import Conexao

def sql_materiais():
    sql = """
    select cod_material
    ,cod_familia
    ,cod_grupomaterial
    ,descricao
    ,cod_unidade

    from material.MATERIAL

    where 1=1
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('MATERIAIS.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'COD_MATERIAL')
    ws.write(r, c+1, 'COD_FAMILIA')
    ws.write(r, c+2, 'COD_GRUPOMATERIAL')
    ws.write(r, c+3, 'DESCRICAO')
    ws.write(r, c+4, 'COD_UNIDADE')

    for col1, col2, col3, col4, col5 in linhas:
        r += 1
        ws.write(r, c, col1)
        ws.write(r, c+1, col2)
        ws.write(r, c+2, col3)
        ws.write(r, c+3, col4)
        ws.write(r, c+4, col5)

    wb.close()
    cursor.close()

    con.close()

    print("Database Materiais Atualizada!")

