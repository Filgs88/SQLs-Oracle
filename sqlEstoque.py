import xlsxwriter
from ConexaoOracle import Conexao

def sql_estoque_geral():
    sql = """
    with tabela_quanti1 as(select cod_material, cod_almoxarifado, ano, mes, quantidade

    from (select cod_material, 
    cod_almoxarifado, 
    ano,
    mes, 
    quantidade, 
    max(ANO) over (partition by cod_material) max_ano
    from material.estoque)

    where ano = max_ano)

    , tabela_quantifinal as(select cod_material, cod_almoxarifado, ano, mes, quantidade

    from (select cod_material, 
    cod_almoxarifado, 
    ano,
    mes, 
    quantidade, 
    max(mes) over (partition by cod_material) max_mes

    from tabela_quanti1)

    where mes = max_mes
    and cod_almoxarifado = 8)

    , tabela_medio1 as(select cod_material, ano, mes, custo_medio

    from (select cod_material, 
    ano,
    mes, 
    custo_medio,
    max(ANO) over (partition by cod_material) max_ano
    from material.customedio)

    where ano = max_ano)

    , tabela_mediofinal as(select cod_material, ano, mes, custo_medio

    from (select cod_material, 
    ano,
    mes, 
    custo_medio, 
    max(mes) over (partition by cod_material) max_mes

    from tabela_medio1)

    where mes = max_mes)


    select material.cod_material
    ,material.cod_familia
    ,material.cod_grupomaterial
    ,material.descricao
    ,coalesce(tabela_quantifinal.quantidade,0) as quantidade
    ,material.cod_unidade
    ,coalesce(tabela_mediofinal.custo_medio,0) as custo_medio

    from tabela_quantifinal , tabela_mediofinal, material.material

    where tabela_quantifinal.cod_material (+)= tabela_mediofinal.cod_material
    and material.cod_material (+)= tabela_mediofinal.cod_material
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('MateiralEstoqueGeral.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'COD_MATERIAL')
    ws.write(r, c+1, 'COD_FAMILIA')
    ws.write(r, c+2, 'COD_GRUPOMATERIAL')
    ws.write(r, c+3, 'DESCRICAO')
    ws.write(r, c+4, 'QUANTIDADE')
    ws.write(r, c+5, 'COD_UNIDADE')
    ws.write(r, c+6, 'CUSTO_MEDIO')


    for col1, col2, col3, col4, col5, col6, col7 in linhas:
        r += 1
        ws.write(r, c, col1)
        ws.write(r, c+1, col2)
        ws.write(r, c+2, col3)
        ws.write(r, c+3, col4)
        ws.write(r, c+4, col5)
        ws.write(r, c+5, col6)
        ws.write(r, c+6, col7)
        
    wb.close()
    cursor.close()
    con.close()

    print("Database Estoque Geral Atualizada!")
