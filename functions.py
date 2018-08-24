import win32com.client
import win32com
import os
import pickle
import time
import re
import numpy as np
np.set_printoptions(threshold=np.nan)

SUBJECT_INFRACAST_MONITOR = 'GEM Production Alerts'
GATEWAY = 'Gateway'
TOTAL_ROW = 7

def get_subject_email(message):
    return message.Subject

def get_body_email(message):
    return message.body

def get_last_email(message):
    return message.GetLast()

def get_number_gateway(message):
    message_aux = get_last_email(message)
    subject = message_aux.Subject
    body = message_aux.body
    LIST_NUMBER_GATEWAY = []
    number_gateway = ''
    FIND_GATEWAY = False

    if subject == SUBJECT_INFRACAST_MONITOR:
        body = body.split(' ')
        for word in body:
            if FIND_GATEWAY == True:
                number_gateway = word
                len_temp = len(number_gateway)
                number_gateway = number_gateway[:len_temp - 1]
                FIND_GATEWAY = False
                LIST_NUMBER_GATEWAY.append(number_gateway)
            if str(word) == GATEWAY:
                FIND_GATEWAY = True

    return LIST_NUMBER_GATEWAY

'''
def get_name_country(message):
    message_aux = get_last_email(message)
    subject = message_aux.Subject
    body = message_aux.body
    LIST_NAME_GOUNTRY = []
    name_country = ''
    FIND_GATEWAY = False
    FIND_NAME_GOUNTRY = False

    if subject == SUBJECT_INFRACAST_MONITOR:
        body = body.split(' ')
        for word in body:
            if FIND_GATEWAY == True and FIND_NAME_GOUNTRY == True:
                name_country = word
                FIND_GATEWAY = False
                FIND_NAME_GOUNTRY = False
                LIST_NAME_GOUNTRY.append(name_country)
            if FIND_GATEWAY == True:
                FIND_NAME_GOUNTRY = True
            if str(word) == GATEWAY:
                FIND_GATEWAY = True

    return LIST_NAME_GOUNTRY

def get_country_code(message):
    message_aux = get_last_email(message)
    subject = message_aux.Subject
    body = message_aux.body
    LIST_GOUNTRY_CODE = []
    country_code = ''
    FIND_GATEWAY = False
    FIND_NAME_GOUNTRY = False
    FIND_GOUNTRY_CODE = False

    if subject == SUBJECT_INFRACAST_MONITOR:
        body = body.split(' ')
        for word in body:
            if FIND_GATEWAY == True and FIND_NAME_GOUNTRY == True and FIND_GOUNTRY_CODE == True:
                country_code = word
                FIND_GATEWAY = False
                FIND_NAME_GOUNTRY = False
                FIND_GOUNTRY_CODE = False
                LIST_GOUNTRY_CODE.append(country_code)
            if FIND_NAME_GOUNTRY == True:
                FIND_GOUNTRY_CODE = True
            if FIND_GATEWAY == True:
                FIND_NAME_GOUNTRY = True
            if str(word) == GATEWAY:
                FIND_GATEWAY = True

    return LIST_GOUNTRY_CODE
'''

def get_country_code(message):
    message_aux = get_last_email(message)
    message_aux = get_body_email(message_aux)
    pattern_country_code = re.compile('(\( *[0-9]+ *\))')
    return pattern_country_code.findall(str(message_aux))

def get_normalize_country_code(list_country_code):
    LIST_NORMALIZE_COUNTRY_CODE = []
    for country_code in list_country_code:
        country_code_aux = country_code.strip('\(')
        country_normalize_country = country_code_aux.strip('\)')
        LIST_NORMALIZE_COUNTRY_CODE.append(country_normalize_country)
    return LIST_NORMALIZE_COUNTRY_CODE

def get_name_country(list_country_code):
    LIST_NAME_COUNTRY = []
    for country_code in list_country_code:
        if country_code == '507':
            LIST_NAME_COUNTRY.append('Panama')
        elif country_code == '505':
            LIST_NAME_COUNTRY.append('Nicaragua')
        elif country_code == '503':
            LIST_NAME_COUNTRY.append('EL Salvador')
        elif country_code == '90':
            LIST_NAME_COUNTRY.append('Turkey')
        elif country_code == '593':
            LIST_NAME_COUNTRY.append('Ecuador')
        elif country_code == '34':
            LIST_NAME_COUNTRY.append('España')
        elif country_code == '351':
            LIST_NAME_COUNTRY.append('Portugal')
        elif country_code == '33':
            LIST_NAME_COUNTRY.append('France')
        elif country_code == '56':
            LIST_NAME_COUNTRY.append('Chile')
        elif country_code == '44':
            LIST_NAME_COUNTRY.append('United Kingdom')
        else:
            print ("Pais nuevo añadir!!!!")
    return LIST_NAME_COUNTRY

def get_high_rate(message):
    message_aux = get_last_email(message)
    message_aux = get_body_email(message_aux)
    pattern_high_rate = re.compile('(high *[a-z]+ *rate)')
    return pattern_high_rate.findall(str(message_aux))

def get_percent(message):
    message_aux = get_last_email(message)
    message_aux = get_body_email(message_aux)
    pattern_percent = re.compile('(\( *[0-9]+ *%)')
    return pattern_percent.findall(str(message_aux))

def get_normalize_percent(list_percent):
    list_normalize_percent = []
    for percent in list_percent:
        percent_aux = percent.strip('\(')
        percent_normalize = percent_aux.strip('%')
        list_normalize_percent.append(percent_normalize)
    return list_normalize_percent

def get_normalize_percent(list_percent):
    list_normalize_percent = []
    for percent in list_percent:
        percent_aux = percent.strip('\(')
        percent_normalize = percent_aux.strip('%')
        list_normalize_percent.append(percent_normalize)
    return list_normalize_percent

def get_out(message):
    message_aux = get_last_email(message)
    message_aux = get_body_email(message_aux)
    out_percent = re.compile('(\), *[0-9]+ *out)')
    return out_percent.findall(str(message_aux))

def get_normalize_out(list_out):
    list_normalize_out = []
    for out in list_out:
        out_aux = out.strip('\),')
        out_normalize = out_aux.strip('out')
        out_normalize = out_normalize.strip()
        list_normalize_out.append(out_normalize)
    return list_normalize_out

def get_messages(message):
    message_aux = get_last_email(message)
    message_aux = get_body_email(message_aux)
    messages_percent = re.compile('(of *[0-9]+ *messages)')
    return messages_percent.findall(str(message_aux))

def get_normalize_messages(list_messages):
    list_normalize_messages = []
    for messages in list_messages:
        messages_aux = messages.strip('of')
        messages_normalize = messages_aux.strip('messages')
        messages_normalize = messages_normalize.strip()
        list_normalize_messages.append(messages_normalize)
    return list_normalize_messages

def count_element_list(list):
    count_total = 0
    for element in list:
        count_total = count_total + 1
    return count_total

def concat_list_final(count_element_total, list_number_gateway, list_normalize_country_code, list_name_country, list_high_rate,
                 list_normalize_percent , list_normalize_out, list_normalize_messages):
    list_final= []
    list_final[0] = 9
    '''
    matriz = np.empty((count_element_total,TOTAL_ROW))
    total_filas = 0
    total_columnas = 0

    for i in list_number_gateway:
        print (i)
        matriz[total_filas][total_columnas] = list_number_gateway[total_filas]
        total_filas = total_filas + 1

    total_columnas = total_columnas + 1
    total_filas = 0

    for i in list_number_gateway:
        print (i)
        matriz[total_filas][total_columnas] = list_normalize_country_code[total_filas]
        total_filas = total_filas + 1

    total_columnas = total_columnas + 1
    total_filas = 0

    for i in list_number_gateway:
        print (i)
        matriz[total_filas][total_columnas] = list_name_country[total_filas]
        total_filas = total_filas + 1

    print (matriz)
    '''