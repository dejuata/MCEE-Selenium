# -*- coding: utf-8 -*-
import re
import random

from datetime import datetime

def validar_correo(value):
  email = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', value)
  if email == None:
    return False
  else:
    return True

def convertir_tipo_documento(value):
  arr = [
    'Registro civil (RC)',
    'Tarjeta de identidad (T.I.)',
    'Cédula de ciudadanía (CC)'
  ]
  return arr[value-1]

def convertir_sexo(value):
  if value == 1:
    return 'Mujer'
  if value == 2:
    return 'Hombre'

def convertir_etnia(value):
  arr = [
    'Negro, mulato(a), Afrodescendiente, Afrocolombiano(a)',
    'Raizal del archipielago de San Andres, Providencia, Santa Catalina',
    'Palenquero(a) de San Basilio',
    'Indígena',
    'Gitano(a) o Rom',
    'Ninguno',
    'No sabe',
    'No responde'
  ]
  return arr[int(value)-1]

def convertir_si_no(value):
  if value == 1:
    return 'Si'
  if value == 2:
    return 'No'

def convertir_tipo_discapacidad(value):
  arr = [
    'Auditiva',
    'Cognitiva',
    'Física',
    'Psicosocial/Mental',
    'Visual',
    'Visual'
  ]
  return arr[int(value)-1]

def generar_fecha(value):
  arr = value.split('/')
  day = random.randint(1,27)
  month = random.randint(1,12)
  year = arr[2]
  fecha = "{}/{}/{}".format(day, month, year)
  return datetime.strptime(fecha, '%d/%m/%y')



def validar_tipo_entero(value):
  try:
    int(value)
    return True
  except:
    return False

def validar_tamano(value):
  return len(str(value)) > 10

