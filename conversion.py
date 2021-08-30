#!/usr/bin/env python3
# -*- encoding: utf-8 -*-


from openpyxl import load_workbook

import psycopg2
import config
import re

ver = "1.0"
copyleft = "(c) 2021 Viktor"
comment = "Скрипт для конвертации таблиц Excel в базу даннух Postgres"

print(f"{ver} \n{copyleft}\n{comment}\n\n")

sourseFileUserUNEP = 'D:\\project\\conversion Excel in Postgres\\!initial data\\UNEP_users.xlsx'
sourseFileUserUKEP = 'D:\\project\\conversion Excel in Postgres\\!initial data\\UKEP_users.xlsx'
sourseFileSertUNEP = 'D:\\project\\conversion Excel in Postgres\\!initial data\\UNEP_certs.xlsx'
sourseFileSertUKEP = 'D:\\project\\conversion Excel in Postgres\\!initial data\\UKEP_certs.xlsx'
file_error = 'D:\\project\\conversion Excel in Postgres\\!initial data\\data_error.txt'

userList = {}
errorList = []  #

db_config = config.db_config

conn = psycopg2.connect(**db_config)
cursor = conn.cursor()

print("Информация о сервере PostgreSQL")
print(conn.get_dsn_parameters(), "\n")

cursor.close()
conn.close()

def createTableUCUser():
    conn = psycopg2.connect(**db_config)
    cursor = conn.cursor()

    try:
        cursor.execute('''CREATE TABLE UCUser
            (ID VARCHAR(50) PRIMARY KEY UNIQUE,
            UPN VARCHAR(50) NOT NULL);''')

        print('Таблица UCUser создана')

    except psycopg2.errors.DuplicateTable:
        print ('Таблица UCUser существует')

    conn.commit()
    cursor.close()
    conn.close()

def createTableCertType():
    conn = psycopg2.connect(**db_config)
    cursor = conn.cursor()

    try:
        cursor.execute('''CREATE TABLE CertificateType
            (ID SERIAL PRIMARY KEY,
            certificateTypeName CHAR(20) UNIQUE,
            CertificateTypeDescription VARCHAR(100));''')

        print('Таблица CertificateType создана')

    except psycopg2.errors.DuplicateTable:
        print ('Таблица CertificateType существует')

    conn.commit()

    _SQL = """INSERT INTO public.CertificateType (certificateTypeName) VALUES ('UNEP')"""
    cursor.execute(_SQL)

    _SQL = """INSERT INTO public.CertificateType (certificateTypeName) VALUES ('UKEP')"""
    cursor.execute(_SQL)
    conn.commit()
    cursor.close()
    conn.close()

    print('Tablen ')

#
# try:
#     cursor.execute('''CREATE TABLE CertificateUNEP
#         (ID SERIAL PRIMARY KEY,
#         CertificateType int REFERENCES CertificateType(ID) ON DELETE CASCADE,
#         Thumbprint VARCHAR(50) no NULL,
#         validBefore DATE no NULL,
#         validAfter DATE no NULL,
#         IsDefault boolean,
#         UCUser VARCHAR(50) REFERENCES UCUser(ID) ON DELETE CASCADE
#         );''')
#
#     print('Таблица Certificate создана')
#
# except psycopg2.errors.DuplicateTable:
#     print ('Таблица Certificate существует')
#
#
# '''
# 1. CertificateType: id(long), certificateTypeName (String), CertificateTypeDescription(string)
# 2. UCUser: id, upn(String)
# 3.    Certificate: id, CertificateType(fk), serial(String), valid Before(Date), validAfter(Date),
#                                                               IsDefault(Boolean), uCUser(fk)
# '''
#
# conn.commit()
# conn.close()

def readUsersUnep():
    wb = load_workbook(sourseFileUserUNEP)
    ws = wb['Лист1']
    i = 0

    for row in range(2, ws.max_row + 1):
        if ws["A" + str(row)].value:
            userID = ws["A" + str(row)].value  # id из файла Эксель
            providerKey = str(ws["B" + str(row)].value)  # Логин из файла Эксель
            i += 1

            if userID not in userList:
                userList[userID] = providerKey
                print(f'--- Записываем значение {i} ----{userID} {providerKey}')

            elif userID in userList and '@' in providerKey:
                print(f'--- Записываем значение {i} ----{userID} {providerKey}')
                userList[userID] = providerKey
'''
    with open('userList.txt', 'w') as f:
        for item, key in userList.items():  # Вывод на печать справочника
            print(item, key)
            f.write(f'{item}, {key}\n')
'''

'''
Функция проверяет всели данные из таблицы занесены в БД            
def test():
    wb = load_workbook(sourseFileUserUNEP)
    ws = wb['Лист1']
    i = 0
    print("Начало теста")

    for row in range(2, ws.max_row+1):
        if ws["A"+str(row)].value:
            userID = ws["A"+str(row)].value # id из файла Эксель
            providerKey = str(ws["B"+str(row)].value) # Логин из файла Эксель

            if userList[userID].split('@')[0] != providerKey.split('@')[0]:
                print (f'Ошибка {userList[userID]}/{providerKey}')
'''

'''
Код для проверки соотношения ID только одному User. Если один ID оответствует разным Логинам пишем в файл error_UNEP.txt

            if userID in userList and userList[userID].split('@')[0] != providerKey.split('@')[0]:
                with open('error_UNEP.txt', 'a') as f:
                    f.write(f'id = {userID} {userList[userID].split("@")[0]} - {providerKey.split("@")[0]}\n')
            else:
                userList[userID] = providerKey
'''

'''
код проверки сколько логинов Users без @ 
    # i = 0
    # with open ('error_UNEP.txt', 'a') as f:
    #     for item in errorList:
    #         if item not in userList:
    #             i +=1
    #             f.write(f'{i} ID без @ - {item}\n')
    #
    #     print (i)
'''

def saveTableUserUNEP():
    i = 0
    conn = psycopg2.connect(**db_config)
    cursor = conn.cursor()
    for UserID , ProviderKey in userList.items():
        i +=1
        print(f'---Пишим в БД {i} значение')
        _SQL = """INSERT INTO public.UCUser(ID, UPN)
            VALUES ('%(UserID)s', '%(ProviderKey)s')
            """%{'UserID':UserID, 'ProviderKey':ProviderKey}

        cursor.execute(_SQL)
        conn.commit()

    print('Table UCUser write in Postgres')


# createTableUCUser()
# readUsersUnep()
# saveTableUserUNEP()
createTableCertType()

