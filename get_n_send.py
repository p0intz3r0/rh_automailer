import pymysql as mysql
import win32com.client as w32clt
import sys, os

def connect(conf_list):
    conn=mysql.connect(host=conf_list[0],
                       user=conf_list[1],
                       password=conf_list[2],
                       db=conf_list[3]
                       )
    try:
        with conn.cursor() as cursor:
            sql = "SELECT email FROM c_user WHERE c_user.company='INTERNAL' ORDER BY lastname ASC;"
            cursor.execute(sql)
            mail_adr = cursor.fetchall()
            list_mail = []
            for i in mail_adr:
                list_mail.append((str(i).strip("'(),")))
            print(list_mail)
            mail_to_list(list_mail)


    except ConnectionError or ConnectionRefusedError  :
        print('Erreur de connection. Verifiez votre mot de passe')
        conn.close()

    finally:
        conn.close()

def read_credentials():
    conf_file = open(os.path.dirname(os.path.abspath(__file__))+ '\\' + 'server.conf', 'r')
    conf_list = []
    for line in conf_file:
        conf_list.append(str(line).strip('\n'))
    print(conf_list)
    connect(conf_list)

def mail_to_list(list_mail):
    app = w32clt.Dispatch("Outlook.Application")
    msg = app.createItem(0)
    msg.BCC =  str(list_mail).strip("'][").replace(',',';').replace("'", "")
    msg.Body = "This is a TEST"
    msg.Save()


if __name__ == '__main__':
    read_credentials()