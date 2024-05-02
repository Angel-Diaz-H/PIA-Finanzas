#--------------------------------------MÓDULOS.
import os
import pandas as pd
import sqlite3
from sqlite3 import Error
import sys
from unidecode import unidecode
import re
from tabulate import tabulate
from colorama import Fore, Style
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#--------------------------------------FUNCIONES DE IMPRESIÓN CON ESTILOS.
def limpiar_consola():
    os.system('cls' if os.name == 'nt' else 'clear')

def aviso(mensaje, longitud):
    mensaje = mensaje.upper()
    print(f"\n{Style.BRIGHT}{guiones(longitud)} {mensaje} {guiones(longitud)}{Style.RESET_ALL}\n")

def avisoGreen(mensaje, longitud):
    mensaje = mensaje.upper()
    print(f"\n{Fore.GREEN}{Style.BRIGHT}{guiones(longitud)} {mensaje} {guiones(longitud)}{Style.RESET_ALL}\n")
    
def avisoRed(mensaje, longitud):
    mensaje = mensaje.upper()
    print(f"\n{Fore.RED}{Style.BRIGHT}{guiones(longitud)} {mensaje} {guiones(longitud)}{Style.RESET_ALL}\n")

def avisoBlue(mensaje, longitud):
    mensaje = mensaje.upper()
    print(f"\n{Fore.BLUE}{Style.BRIGHT}{guiones(longitud)} {mensaje} {guiones(longitud)}{Style.RESET_ALL}\n")


def printNegrita(mensaje):
    print(f"{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def inputNegrita(mensaje):
    return input(f"{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")


def printGreen(mensaje):
    print(f"{Fore.GREEN}{mensaje}{Style.RESET_ALL}")

def printRed(mensaje):
    print(f"{Fore.RED}{mensaje}{Style.RESET_ALL}")

def printBlue(mensaje):
    print(f"{Fore.BLUE}{mensaje}{Style.RESET_ALL}")

def printCyan(mensaje):
    print(f"{Fore.CYAN}{mensaje}{Style.RESET_ALL}")


def printGreenNegrita(mensaje):
    print(f"{Fore.GREEN}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def printRedNegrita(mensaje):
    print(f"{Fore.RED}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def printBlueNegrita(mensaje):
    print(f"{Fore.BLUE}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def printCyanNegrita(mensaje):
    print(f"{Fore.CYAN}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")


def inputGreen(mensaje):
    return input(f"{Fore.GREEN}{mensaje}{Style.RESET_ALL}")

def inputRed(mensaje):
    return input(f"{Fore.RED}{mensaje}{Style.RESET_ALL}")

def inputBlue(mensaje):
    return input(f"{Fore.BLUE}{mensaje}{Style.RESET_ALL}")

def inputCyan(mensaje):
    return input(f"{Fore.CYAN}{mensaje}{Style.RESET_ALL}")

def inputGreenNegrita(mensaje):
    return input(f"{Fore.GREEN}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def inputRedNegrita(mensaje):
    return input(f"{Fore.RED}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def inputBlueNegrita(mensaje):
    return input(f"{Fore.BLUE}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")

def inputCyanNegrita(mensaje):
    return input(f"{Fore.CYAN}{Style.BRIGHT}{mensaje}{Style.RESET_ALL}")
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#--------------------------------------FUNCIONES GENERALES.
def guiones(longitud):
    return '-' * longitud

def respuestaSINO():
    while True:
        respuesta = darFormatoATexto(inputNegrita('\tRespuesta: '))
        if respuesta == 'SI' or respuesta == 'NO':
            break
        else:
            printNegrita('\n\tIngrese una respuesta válida (Sí/No).')
    return respuesta

def darFormatoATexto(texto):
    texto = unidecode(texto).strip().upper()
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def indicarEnter():
    inputBlueNegrita("\n\nDe clic en Enter para continuar.")

#--------------------------------------FUNCIONES AUXILIARES. 
def mostrarOpcionesDeMenu(listaActual):
    print('Ingrese el número de la opción que desee realizar:')
    print(tabulate(listaActual, headers = 'firstrow', tablefmt = 'pretty'))

def validarOpcionesNumericas(opcionMin, opcionMax):
    while True:
        try:
            opcionValida = int(inputCyan("Opción: "))
            if opcionValida >= opcionMin and opcionValida <= opcionMax:
                return opcionValida
            else:
                printNegrita(f"\nIngrese un número válido ({opcionMin}-{opcionMax}).")
        except ValueError:
            printNegrita(f"\nIngrese un número válido ({opcionMin}-{opcionMax}).")

def contarCantidadOpcionesDeMenu(listaActual):
    cantidad = len(listaActual) - 1
    return cantidad

def mostrarYValidarMenu(ubicacion, opcion, lista):
    mostrarTitulo(ubicacion, True)  
    mostrarOpcionesDeMenu(lista)
    opcion = validarOpcionesNumericas(1, contarCantidadOpcionesDeMenu(lista))
    limpiar_consola() 
    if not opcion == contarCantidadOpcionesDeMenu(lista):
            ubicacion.append(lista[opcion][1])
    return opcion, ubicacion

def mostrarTitulo(ubicacion, esMenu = False):
    printNegrita(f'Ubicación: {" / ".join(ubicacion)}')
    if esMenu == True:
        avisoBlue(f'MENÚ {ubicacion[-1]}', 20)
    else:
        avisoBlue(f'{ubicacion[-1]}', 20)
        return
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#--------------------------------------IMPORTAR EL EXCEL.
aviso("Aviso:", 0)
try:
    df = pd.read_excel('Antiguedad_de_saldos.xlsx', engine = 'openpyxl')
    printGreenNegrita("El archivo Excel se ha leído correctamente.\n")
except Error as e:
        printRedNegrita(f'Se produjo el siguiente error: {e}')
except Exception:
        printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

#Copia el df a otro sin nombres duplicados.
try:
    df_copia1 = df[['Nombre']].copy()
    df_copia1 = df_copia1.drop_duplicates().reset_index(drop = True)
except Error as e:
        printRedNegrita(f'Se produjo el siguiente error: {e}')
except Exception:
        printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

#--------------------------------------CREACIÓN DE TABLAS SQL (SOLO SI NO EXISTEN).
def creacion_tablas():
    try:
        with sqlite3.connect('CuentasPorCobrar.db') as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS Clientes \
                            (CLAVE_CLIENTE INTEGER PRIMARY KEY NOT NULL,\
                            NOMBRECLIENTE TEXT NOT NULL);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS CuentasPorCobrar \
                            (CLAVE_GUIA INTEGER PRIMARY KEY NOT NULL, \
                            DIAS INTEGER NOT NULL,\
                            FECHA timestamp NOT NULL, \
                            TOTAL INTEGER NOT NULL, \
                            CLAVE_CLIENTE INTEGER NOT NULL,\
                            FOREIGN KEY (CLAVE_CLIENTE) REFERENCES CLIENTES(CLAVE_CLIENTE));")
            printGreenNegrita("Las tablas SQL se han cargado correctamente\n")
    except Error as e:
        printRedNegrita(f'Se produjo el siguiente error: {e}')
    except Exception:
        printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
    finally:
        conn.close()
creacion_tablas()

#--------------------------------------INSERTAR REGISTROS DE EXCEL A LA TABLA CLIENTES (SOLO SI NO EXISTEN).
try:
    with sqlite3.connect('CuentasPorCobrar.db') as conn:
        mi_cursor = conn.cursor()
        for nombre_cliente in df_copia1['Nombre']:
            mi_cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE CLAVE_CLIENTE=?", (nombre_cliente,))
            existencia = mi_cursor.fetchone()[0]
            if existencia == 0:
                mi_cursor.execute("INSERT INTO CLIENTES(NOMBRECLIENTE) VALUES (?)", (nombre_cliente,))
        printGreenNegrita("La información de los clientes se han registrado satisfactoriamente.\n")
except Error as e:
    printRedNegrita(f'Se produjo el siguiente error: {e}')
except Exception:
    printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
finally:                                    
    conn.close()

#INSERTAR REGISTROS DE EXCEL A LA TABLA CUENTASPORCOBRAR (SOLO SI NO EXISTEN Y SI CLAVE_CLIENTE EXISTE EN TABLAS CLIENTES).
try:
    with sqlite3.connect('CuentasPorCobrar.db') as conn:
        mi_cursor = conn.cursor()
        for index, row in df.iterrows():
            mi_cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE CLAVE_CLIENTE=?", (row['Cliente'],))
            existencia2 = mi_cursor.fetchone()[0]

            mi_cursor.execute("SELECT COUNT(*) FROM CuentasPorCobrar WHERE CLAVE_GUIA=?", (row['Guía'],))
            existencia3 = mi_cursor.fetchone()[0]

            if existencia2 > 0 and existencia3 == 0:
                fecha = row['Fecha'].strftime('%Y-%m-%d')  #Convierte la fecha a una cadena de texto.
                mi_cursor.execute("INSERT INTO CuentasPorCobrar(CLAVE_GUIA, DIAS, FECHA, TOTAL, CLAVE_CLIENTE) VALUES (?, ?, ?, ?, ?)", 
                                  (row['Guía'], row['Días '], fecha, row['Total'], row['Cliente']))
        printGreenNegrita("La información de cuentas por cobrar se ha registrado correctamente.")
except Error as e:
    printRedNegrita(f'Se produjo el siguiente error: {e}')
except Exception:
    printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
finally:                                    
    conn.close()

indicarEnter()
limpiar_consola()
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#--------------------------------------LISTAS DE MENÚS.
lmenu_principal = [('Opción', 'Descripción'),
              (1, 'Clientes'),
              (2, 'Cuentas por pagar'),
              (3, 'Salir')]
lmenu_clientes = [('Opción', 'Descripción'),
              (1, 'Registrar clientes'),
              (2, 'Eliminar clientes'),
              (3, 'Recuperar clientes'),
              (4, 'Mostrar clientes'),
              (5, 'Volver al menú principal')]
lmenu_cuentaPorPagar = [('Opción', 'Descripción'),
              (1, 'Registrar cuentras por pagar'),
              (2, 'Eliminar cuentas por pagar'),
              (3, 'Recuperar cuentas por pagar'),
              (4, 'Mostrar cuentas por pagar'),
              (5, 'Análisis de cuentas por pagar'),
              (6, 'Volver al menú principal')]
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#
#--------------------------------------1. MENÚ PRINCIPAL.
def menuPrincipal():
    opcion = 0

    while True:    
        ubicacion = ['Menú principal']
        if opcion == 0:
            avisoBlue('CUENTAS POR COBRAR', 20)
            mostrarOpcionesDeMenu(lmenu_principal)

            opcion = validarOpcionesNumericas(1, contarCantidadOpcionesDeMenu(lmenu_principal))

        if not opcion == contarCantidadOpcionesDeMenu(lmenu_principal):
            ubicacion.append(lmenu_principal[opcion][1])
            limpiar_consola()

        if opcion == 1:
            menuClientes(ubicacion)
        elif opcion == 2:
            menuCuentasPorPagar(ubicacion)
        else:
            printBlueNegrita('\n¿Está seguro que desea salir? (Sí/No)')
            respuesta = respuestaSINO()
            if respuesta == 'SI':
                avisoGreen("Archivo cerrado correctamente.", 25)
                avisoGreen("Gracias por usar nuestro sistema, hasta la próxima.", 15)
                break
            else:
                limpiar_consola()
                opcion = 0

        opcion = 0
        continue

#--------------------------------------2.1. MENÚ CLIENTES.
def menuClientes(ubicacion):
    ubicacionOriginal = ubicacion.copy()
    opcion = 0

    while True:    
        ubicacion = ubicacionOriginal.copy()
        if opcion == 0:
            opcion, ubicacion = mostrarYValidarMenu(ubicacion, opcion, lmenu_clientes)

        if opcion == 1:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 2:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 3:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 4:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        else:
            break

        opcion = 0
        limpiar_consola()
        continue

#--------------------------------------2.2. MENÚ CUENTAS POR PAGAR.
def menuCuentasPorPagar(ubicacion):
    ubicacionOriginal = ubicacion.copy()
    opcion = 0

    while True:    
        ubicacion = ubicacionOriginal.copy()
        if opcion == 0:
            opcion, ubicacion = mostrarYValidarMenu(ubicacion, opcion, lmenu_cuentaPorPagar)

        if opcion == 1:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 2:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 3:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 4:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        elif opcion == 5:
            mostrarTitulo(ubicacion)
            #función_correspondiente()
        else:
            break

        opcion = 0
        limpiar_consola()
        continue

menuPrincipal()
