import pyodbc

class Conexao:
    def conection():
        con = pyodbc.connect(
        'Driver={Driver};'
        'dbq=host:port/service_name;'
        'Uid=User ID;'
        'Pwd=Password;'
        )
        return con
        
    def close(con):
        con.close()
