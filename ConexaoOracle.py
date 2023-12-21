import pyodbc

class Conexao:
    def conection():
        con = pyodbc.connect(
        'Driver={Oracle em OraClient10g_home1};'
        'dbq=192.168.1.12:1521/csorcl;'
        'Uid=fatimaagro;'
        'Pwd=A#QNK2bdJh8US;'
        )
        return con
        
    def close(con):
        con.close()