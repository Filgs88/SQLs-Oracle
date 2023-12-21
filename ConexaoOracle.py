import pyodbc

class Conexao:
    def conection():
        con = pyodbc.connect(
        'Driver={Oracle em OraClient10g_home1};'
        'dbq=host:port/service_name;'
        'Uid=User ID;'
        'Pwd=Password;'
        )
        return con
        
    def close(con):
        con.close()
