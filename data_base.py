import mysql.connector as mariadb


def insert_mysql(count_element_total_split, list_final):
    MARIADB_CONNECTION = mariadb.connect(user='root', password='', database='infracast_monitor')
    CURSOR = MARIADB_CONNECTION.cursor()
    count_field = 0
    country = ''
    insert = False
    for i in range(0, count_element_total_split):
        if count_field == 0:
            date = list_final[i]
            count_field = count_field + 1
        elif count_field == 1:
            gateway = list_final[i]
            count_field = count_field + 1
            print("gateway: ", gateway)
        elif count_field == 2:
            country_code = int(list_final[i])
            count_field = count_field + 1
        elif count_field == 3:
            country = str(list_final[i])
            count_field = count_field + 1
        elif count_field == 4:
            high_rate = list_final[i]
            count_field = count_field + 1
        elif count_field == 5:
            percent = int(list_final[i])
            count_field = count_field + 1
        elif count_field == 6:
            affected_messages = int(list_final[i])
            count_field = count_field + 1
        elif count_field == 7:
            total_messages = int(list_final[i])
            count_field = 0
            insert = True
        if insert == True:
            CURSOR.execute("INSERT INTO alarmas_infracast (id, date, gateway, country_code, country, high_rate, percent, affected_messages, total_messages) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)", (None, date, gateway, country_code, country, high_rate, percent, affected_messages, total_messages))
            MARIADB_CONNECTION.commit()
            print(CURSOR.rowcount, "record inserted.")
            insert = False