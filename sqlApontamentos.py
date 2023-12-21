import xlsxwriter
from ConexaoOracle import Conexao

def sql_apontamentos():
    sql = """
    select data_inicio
    ,cod_funcionario
    ,coalesce(tothoras_trabalhadas,0) as tothoras_trabalhadas
    ,ano_ordem_servico
    ,numero_ordem_servico
    ,codigo_tarefa
    ,codigo_tarefa || '-' || numero_ordem_servico || '-' || ano_ordem_servico as ID

    from industria.item_ordem_servico_apont

    where data_inicio >= to_date('01/08/2023','dd/mm/yyyy')
    """
    
    con = Conexao.conection()

    cursor = con.cursor()
    cursor.execute(sql)
    linhas = cursor.fetchall()

    wb = xlsxwriter.Workbook('Apontamentos.xlsx')
    ws = wb.add_worksheet()

    r = 0
    c = 0

    ws.write(r, c, 'DATA_INICIO')
    ws.write(r, c+1, 'COD_FUNCIONARIO')
    ws.write(r, c+2, 'TOTHORAS_TRABALHADAS')
    ws.write(r, c+3, 'ANO_ORDEM_SERVICO')
    ws.write(r, c+4, 'NUMERO_ORDEM_SERVICO')
    ws.write(r, c+5, 'CODIGO_TAREFA')
    ws.write(r, c+6, 'ID')

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

    print("Database Apontamentos Atualizada!")

