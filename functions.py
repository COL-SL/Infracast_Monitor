import win32com.client
import win32com
import os
import pickle
import time

SUBJECT_INFRACAST_MONITOR = 'GEM Production Alerts'
GATEWAY = 'Gateway'


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

