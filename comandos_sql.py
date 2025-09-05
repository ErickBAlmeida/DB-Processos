import sqlite3

banco = sqlite3.connect(r"C:\Users\ealmeida\Desktop\BarrosBD.db")
cursor = banco.cursor()

def apagar_duplicados():
    sql = """
        DELETE FROM processos
        WHERE rowid NOT IN (
            SELECT rowid
            FROM (
                SELECT MAX(rowid) AS rowid
                FROM processos
                GROUP BY GCPJ, registro, estoque, atraso
                HAVING MAX(meta)
            )
        );"""

    cursor.execute(sql)
    banco.commit()
