import pyodbc
import xlsxwriter


con = pyodbc.connect(
    'Driver={Oracle em OraClient10g_home1};'
    'dbq=192.168.1.12:1521/csorcl;'
    'Uid=fatimaagro;'
    'Pwd=A#QNK2bdJh8US;'
)

cursor = con.cursor()

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
cursor.execute(sql)
linhas = cursor.fetchall()

for linha in linhas:
    print(linha)

cursor.close()
con.close()

print("Planilha Atualizada!")