import MySQLdb


def connect():

    conn = MySQLdb.connect(
        host='localhost',
        port=3306,
        user='root',
        passwd='123456789',
        db='suanfa'
    )

    cur = conn.cursor()


    cur.execute("select * from es_info ")


