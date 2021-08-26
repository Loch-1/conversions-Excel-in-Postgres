#!/usr/bin/env python3
# -*- encoding: utf-8 -*-

#import codecs
#import time
#import xlrd
from openpyxl import load_workbook
#import pandas
import psycopg2
import re

ver="1.0";
copyleft="(c) 2021 Viktor";
comment="Скрипт для конвертации таблиц Excel в базу даннух Postgres"

print(f"{ver} \n{copyleft}\n{comment}\n\n")

db_config = {'database':'cryptobot',
           'user':'postgres',
           'password':'postgres',
           'host':'127.0.0.1'}

conn = psycopg2.connect(**db_config)
cursor = conn.cursor()

try:
    cursor.execute('''CREATE TABLE CertificateType 
        (ID SERIAL PRIMARY KEY,
        certificateTypeName CHAR(50),
        CertificateTypeDescription CHAR(50));''')

    print('Таблица CertificateType создана')

except psycopg2.errors.DuplicateTable:
    print ('Таблица CertificateType существует')

conn.commit()

try:
    cursor.execute('''CREATE TABLE UCUser 
        (ID SERIAL PRIMARY KEY,
        upn VARCHAR(50));''')

    print('Таблица UCUser создана')

except psycopg2.errors.DuplicateTable:
    print ('Таблица UCUser существует')

conn.commit()

try:
    cursor.execute('''CREATE TABLE Certificate 
        (ID SERIAL PRIMARY KEY,        
        CertificateType int REFERENCES CertificateType(ID) ON DELETE CASCADE,        
        serial VARCHAR(50),
        validBefore DATE,
        validAfter DATE,
        IsDefault boolean,        
        uCUser int REFERENCES UCUser(ID) ON DELETE CASCADE
        );''')

    print('Таблица Certificate создана')

except psycopg2.errors.DuplicateTable:
    print ('Таблица Certificate существует')


'''
1. CertificateType: id(long), certificateTypeName (String), CertificateTypeDescription(string)
2. UCUser: id, upn(String)
3.Certificate: id, CertificateType(fk), serial(String), valid Before(Date), validAfter(Date), IsDefault(Boolean), uCUser(fk)
'''

conn.commit()
conn.close()