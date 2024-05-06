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
from datetime import datetime
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
def limpiar_consola():
    os.system('cls' if os.name == 'nt' else 'clear')

def guiones(longitud):
    return '-' * longitud

def respuestaSiYNo():
    while True:
        respuesta = darFormatoATexto(inputNegrita('\tRespuesta: '))
        if respuesta == 'SI':
            return True
        elif respuesta == 'NO':
            return False
        else:
            printNegrita('\n\tIngrese una respuesta válida (Sí/No).')

def darFormatoATexto(texto, sinEspacios = False):
    texto = unidecode(texto).strip().upper()
    if sinEspacios:
        texto = texto.replace(' ', '')
    else:
        texto = re.sub(r'\s+', ' ', texto)
    return texto

def validarSalir(texto):
    texto = darFormatoATexto(str(texto), True)
    if texto == '<REGRESAR>':
        return True
    return False

def indicarEnter():
    inputBlueNegrita("\n\nDe clic en Enter para continuar.")

def validarTextoValido(texto, permitirNumero = True):
    
    if not permitirNumero == True:
        if not texto.replace(' ', '').isalpha():
            printRedNegrita("Ingrese solo letras.\n")
            return True
    else:
        if not re.match("^[A-Za-z0-9]*$", texto.replace(' ', '')):
            printRedNegrita("Ingrese solo letras y números.\n")
            return True

    if len(texto) < 3 or len(texto) > 50:
        printRedNegrita("Ingrese un nombre válido.\n")
        return True
    return False

def mensajeInicialEnFuncionesEspecificas():
    printBlueNegrita("Para regresar al menú anterior ingrese:")
    printGreenNegrita("<Regresar>\n")

def solicitarEnteroOSalir(descripcion):
    while True:
        respuesta = inputNegrita(f"{descripcion}: ")

        if respuesta.isdigit():
            return int(respuesta), False
        elif validarSalir(respuesta):
            return True, True
        else:
            printRedNegrita("Ingrese un número válido.\n")

def solicitarRangoEnteroOSalir(opcionMin, opcionMax):
    while True:
        try:
            opcionValida = int(input("Opción: "))
            if opcionValida >= opcionMin and opcionValida <= opcionMax:
                return opcionValida, False
            elif validarSalir(opcionValida):
                return True, True
            else:
                printRedNegrita(f"Ingrese un número válido ({opcionMin}-{opcionMax}).\n")
        except ValueError:
            printRedNegrita(f"Ingrese un número válido ({opcionMin}-{opcionMax}).\n")
        except Error as e:
            inputRedNegrita(f'Se produjo el siguiente error: {e}')
        except Exception as e:
            inputRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

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
                printRedNegrita(f"Ingrese un número válido ({opcionMin}-{opcionMax}).\n")
        except ValueError:
            printRedNegrita(f"Ingrese un número válido ({opcionMin}-{opcionMax}).\n")

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

#--------------------------------------CREACIÓN DE TABLAS SQL (SOLO SI NO EXISTEN).
def creacion_tablas():
    try:
        with sqlite3.connect('CuentasPorCobrar.db') as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS Clientes \
                            (CLAVE_CLIENTE INTEGER PRIMARY KEY NOT NULL,\
                            NOMBRECLIENTE TEXT NOT NULL, \
                            ESTADOCLIENTE INTEGER NOT NULL);")
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

#--------------------------------------INSERTAR REGISTROS EXCEL A LA TABLA CLIENTES.
def insertar_clientes(df):
    try:
        with sqlite3.connect('CuentasPorCobrar.db') as conn:
            mi_cursor = conn.cursor()
            for i in df['Nombre'].unique():
                mi_cursor.execute("INSERT INTO Clientes (NOMBRECLIENTE, ESTADOCLIENTE) \
                                  SELECT ?, ? WHERE NOT EXISTS(SELECT 1 \
                                  FROM Clientes WHERE NOMBRECLIENTE = ?)", (i, 1, i))
            printGreenNegrita("Los datos se han insertado correctamente en la tabla Clientes.\n")
    except Error as e:
        printRedNegrita(f'Se produjo el siguiente error: {e}')
    except Exception:
        printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
    finally:
        conn.close()

#--------------------------------------INSERTAR REGISTROS EXCEL A LA TABLA CUENTASPORCOBRAR.
def insertar_cuentasporcobrar(df):
    try:
        with sqlite3.connect('CuentasPorCobrar.db') as conn:
            mi_cursor = conn.cursor()
            for i, row in df.iterrows():
                fecha = row['Fecha'].strftime('%Y-%m-%d %H:%M:%S')  # Convertir la fecha a una cadena de texto
                mi_cursor.execute("INSERT INTO CuentasPorCobrar (CLAVE_GUIA, DIAS, FECHA, TOTAL, CLAVE_CLIENTE) \
                                  SELECT ?, ?, ?, ?, ? WHERE NOT EXISTS(SELECT 1 FROM CuentasPorCobrar WHERE CLAVE_GUIA = ?) AND \
                                  EXISTS(SELECT 1 FROM Clientes WHERE CLAVE_CLIENTE = ?)", (row['Guía'], row['Días '], fecha, row['Total'], row['Cliente'], row['Guía'], row['Cliente']))
            printGreenNegrita("Los datos se han insertado correctamente en la tabla CuentasPorCobrar.\n")
    except Error as e:
        printRedNegrita(f'Se produjo el siguiente error: {e}')
    except Exception:
        printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
    finally:
        conn.close()

creacion_tablas()
insertar_clientes(df)
insertar_cuentasporcobrar(df)
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
              (2, 'Suspender clientes'),
              (3, 'Recuperar clientes'),
              (4, 'Mostrar clientes'),
              (5, 'Volver al menú principal')]
lmenu_cuentaPorPagar = [('Opción', 'Descripción'),
              (1, 'Registrar cuentras por pagar'),
              (2, 'Cancelar cuentas por pagar'),
              (3, 'Recuperar cuentas por pagar'),
              (4, 'Mostrar cuentas por pagar'),
              (5, 'Análisis de cuentas por pagar'),
              (6, 'Volver al menú principal')]
lmenu_clientes_orden = [('Opción', 'Orden'),
              (1, 'Por clave'),
              (2, 'Por nombre'),
              (3, 'Por estado')]
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
            printCyanNegrita('\n¿Está seguro que desea salir? (Sí/No)')
            if respuestaSiYNo():
                avisoGreen("Archivo cerrado correctamente.", 25)
                avisoGreen("Gracias por usar nuestro sistema, hasta la próxima.", 15)
                break
            else:
                limpiar_consola()
                opcion = 0

        opcion = 0
        continue

#--------------------------------------1.1. MENÚ CLIENTES.
def menuClientes(ubicacion):
    ubicacionOriginal = ubicacion.copy()
    opcion = 0

    while True:    
        ubicacion = ubicacionOriginal.copy()
        if opcion == 0:
            opcion, ubicacion = mostrarYValidarMenu(ubicacion, opcion, lmenu_clientes)

        if opcion == 1:
            mostrarTitulo(ubicacion)
            registrarClientes()
        elif opcion == 2:
            mostrarTitulo(ubicacion)
            suspenderCliente()
        elif opcion == 3:
            mostrarTitulo(ubicacion)
            recuperarClientes()
        elif opcion == 4:
            mostrarTitulo(ubicacion)
            mostrarClientes()
        else:
            break

        opcion = 0
        limpiar_consola()
        continue

#--------------------------------------1.2. MENÚ CUENTAS POR PAGAR.
def menuCuentasPorPagar(ubicacion):
    ubicacionOriginal = ubicacion.copy()
    opcion = 0

    while True:    
        ubicacion = ubicacionOriginal.copy()
        if opcion == 0:
            opcion, ubicacion = mostrarYValidarMenu(ubicacion, opcion, lmenu_cuentaPorPagar)

        if opcion == 1:
            mostrarTitulo(ubicacion)
            registrarCuentasPorCobrar()
        elif opcion == 2:
            mostrarTitulo(ubicacion)
            cancelarCuentasPorCobrar()
        elif opcion == 3:
            mostrarTitulo(ubicacion)
            recuperarCuentasPorCobrar()
        elif opcion == 4:
            mostrarTitulo(ubicacion)
            mostrarCuentasPorCobrar()
        elif opcion == 5:
            mostrarTitulo(ubicacion)
            analisisDeCuentasPorCobrar()
        else:
            break

        opcion = 0
        limpiar_consola()
        continue
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
#--------------------------------------1.1.1. OPCIÓN REGISTRAR CLIENTES.
def registrarClientes():
    mensajeInicialEnFuncionesEspecificas()
    
    print('Ingrese el nombre del cliente (empresa).')
    while True:
        nombre = darFormatoATexto(inputNegrita('Nombre: '))
        
        if validarSalir(nombre): break
        if validarTextoValido(nombre): continue

        try:
            with sqlite3.connect('CuentasPorCobrar.db') as conn:
                mi_cursor = conn.cursor()
                mi_cursor.execute("SELECT 1 FROM Clientes WHERE NOMBRECLIENTE = ?", (nombre,))
                existencia = mi_cursor.fetchone()
                if existencia: 
                    printRedNegrita("La empresa ya se encuentra registrada.\n")
                    continue
                else:
                    mi_cursor.execute("INSERT INTO Clientes(NOMBRECLIENTE, ESTADOCLIENTE) VALUES (?, ?)", (nombre, 1))
                    printCyanNegrita("\n¿Desea registrar el cliente? Confirme su respuesta (Sí/No).")
                    if respuestaSiYNo():
                        printGreenNegrita("\nRegistro de cliente exitosamente.")
                    else:
                        printBlueNegrita("\nEl cliente no se registró.")
                    indicarEnter()
                    break
        except Error as e:
            printRedNegrita(f'Se produjo el siguiente error: {e}')
        except Exception as e:
            printRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
        finally:
            conn.close()

#--------------------------------------1.1.2. OPCIÓN SUSPENDER CLIENTES.
def suspenderCliente():
    
    while True:
        try:
            with sqlite3.connect('CuentasPorCobrar.db') as conn:
                mi_cursor = conn.cursor()
                mi_cursor.execute("SELECT CLAVE_CLIENTE, NOMBRECLIENTE FROM Clientes WHERE ESTADOCLIENTE = 1")
                existencia = mi_cursor.fetchall()
                if existencia:
                    mensajeInicialEnFuncionesEspecificas()
                    printNegrita("Lista de clientes activos:")
                    print(tabulate(existencia, headers = ["Clave del cliente", "Nombre"], tablefmt = 'pretty'))
                    print("\nIngrese la clave del cliente a suspender.")

                    while True:
                        respuesta = solicitarEnteroOSalir("Clave")
                        if respuesta[1]: break
                        busqueda = any(clave[0] == respuesta[0] for clave in existencia)
                        if busqueda:
                            printCyanNegrita("\n¿Desea suspender el siguiente cliente? (Sí/No)")
                            mi_cursor.execute("SELECT CLAVE_CLIENTE, NOMBRECLIENTE FROM Clientes WHERE CLAVE_CLIENTE = ?", (respuesta[0],))
                            clienteEncontrado = mi_cursor.fetchone()
                            print(tabulate([list(clienteEncontrado)], headers = ["Clave del cliente", "Nombre"], tablefmt = 'pretty'))

                            if respuestaSiYNo():
                                mi_cursor.execute("UPDATE Clientes SET ESTADOCLIENTE = 0 WHERE CLAVE_CLIENTE = ?", (clienteEncontrado[0],))
                                printGreenNegrita("\nCliente suspendido con éxito.")
                            else:
                                printBlueNegrita("\nEl cliente no se suspendió.")
                            indicarEnter()
                            break
                        else:
                            printRedNegrita("Ingrese una clave válida o <regresar> para volver al menú anterior.\n")
                            continue
                else:
                    printBlueNegrita('\nActualmente no se cuenta con clientes activos.')
                    indicarEnter()
                    break
                break
        except Error as e:
            inputRedNegrita(f'Se produjo el siguiente error: {e}')
        except Exception as e:
            inputRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
        finally:
            conn.close()

#--------------------------------------1.1.3. OPCIÓN RECUPERAR CLIENTES.
def recuperarClientes():
    while True:
        try:
            with sqlite3.connect('CuentasPorCobrar.db') as conn:
                mi_cursor = conn.cursor()
                mi_cursor.execute("SELECT CLAVE_CLIENTE, NOMBRECLIENTE FROM Clientes WHERE ESTADOCLIENTE = 0")
                existencia = mi_cursor.fetchall()
                if existencia:
                    mensajeInicialEnFuncionesEspecificas()
                    printNegrita("Lista de clientes suspendidos:")
                    print(tabulate(existencia, headers = ["Clave del cliente", "Nombre"], tablefmt = 'pretty'))
                    print("\nIngrese la clave del cliente a recuperar.")

                    while True:
                        respuesta = solicitarEnteroOSalir("Clave")
                        if respuesta[1]: break
                        busqueda = any(clave[0] == respuesta[0] for clave in existencia)
                        if busqueda:
                            printCyanNegrita("\n¿Desea recuperar el siguiente cliente? (Sí/No)")
                            mi_cursor.execute("SELECT CLAVE_CLIENTE, NOMBRECLIENTE FROM Clientes WHERE CLAVE_CLIENTE = ?", (respuesta[0],))
                            clienteEncontrado = mi_cursor.fetchone()
                            print(tabulate([list(clienteEncontrado)], headers = ["Clave del cliente", "Nombre"], tablefmt = 'pretty'))

                            if respuestaSiYNo():
                                mi_cursor.execute("UPDATE Clientes SET ESTADOCLIENTE = 1 WHERE CLAVE_CLIENTE = ?", (clienteEncontrado[0],))
                                printGreenNegrita("\nCliente recuperado con éxito.")
                            else:
                                printBlueNegrita("\nEl cliente no se recuperó.")
                            indicarEnter()
                            break
                        else:
                            printRedNegrita("Ingrese una clave válida o <regresar> para volver al menú anterior.\n")
                            continue
                else:
                    printBlueNegrita('\nActualmente no se cuenta con clientes suspendidos.')
                    indicarEnter()
                    break
                break
        except Error as e:
            inputRedNegrita(f'Se produjo el siguiente error: {e}')
        except Exception as e:
            inputRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
        finally:
            conn.close()

#--------------------------------------1.1.4. OPCIÓN MOSTRAR CLIENTES.
def mostrarClientes():
    while True:
        try:
            with sqlite3.connect('CuentasPorCobrar.db') as conn:
                mi_cursor = conn.cursor()
                mi_cursor.execute("SELECT CLAVE_CLIENTE, NOMBRECLIENTE, ESTADOCLIENTE FROM Clientes")
                existencia = mi_cursor.fetchall()
                if existencia:
                    mensajeInicialEnFuncionesEspecificas()
                    existenciaImpresa = [(clave, nombre, "Activo" if estado else "Inactivo") for clave, nombre, estado in existencia]
                    mostrarOpcionesDeMenu(lmenu_clientes_orden)
                    respuesta = solicitarRangoEnteroOSalir(1, 3)
                    
                    if respuesta[1]: break
                    printNegrita("\nLista de clientes:")
                    if respuesta[0] == 1:
                        existenciaImpresaOrdenada = sorted(existenciaImpresa, key = lambda x: x[0])
                        print(tabulate(existenciaImpresaOrdenada, headers = ['Clave', 'Nombre', 'Estado'], tablefmt = 'pretty'))
                    elif respuesta[0] == 2:
                        existenciaImpresaOrdenada = sorted(existenciaImpresa, key = lambda x: x[1])
                        print(tabulate(existenciaImpresaOrdenada, headers = ['Clave', 'Nombre', 'Estado'], tablefmt = 'pretty'))
                    else:
                        existenciaImpresaOrdenada = sorted(existenciaImpresa, key = lambda x: x[2])
                        print(tabulate(existenciaImpresaOrdenada, headers = ['Clave', 'Nombre', 'Estado'], tablefmt = 'pretty'))
                    
                    printCyanNegrita("\n¿Desea exportar la información a Excel? (Sí/No)")
                    if respuestaSiYNo():
                        fechaHora = datetime.now().strftime("%d-%m-%Y_%H%M%S")
                        nombre_archivo = f"ListadoClientes_{fechaHora}.xlsx"
                        df = pd.DataFrame(existenciaImpresaOrdenada, columns = ["Clave", "Nombre", "Estado cliente"])
                        df.to_excel(nombre_archivo, index = False)
                        printGreenNegrita("\nInformación exportada exitosamente a Excel.")
                        printBlueNegrita(f"Nombre del archivo:")
                        printNegrita(f"{nombre_archivo}")
                        indicarEnter()
                else:
                    printBlueNegrita('\nActualmente no se cuenta con clientes registrados.')
                    indicarEnter()
                break
        except Error as e:
            inputRedNegrita(f'Se produjo el siguiente error: {e}')
        except Exception as e:
            inputRedNegrita(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
        finally:
            conn.close()
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
#--------------------------------------1.2.1. OPCIÓN REGISTRAR CUENTAS POR COBRAR.
def registrarCuentasPorCobrar():
    inputCyanNegrita("")

#--------------------------------------1.2.2. OPCIÓN CANCELAR CUENTAS POR COBRAR.
def cancelarCuentasPorCobrar():
    inputCyanNegrita("")

#--------------------------------------1.2.3. OPCIÓN RECUPERAR CUENTAS POR COBRAR.
def recuperarCuentasPorCobrar():
    inputCyanNegrita("")

#--------------------------------------1.2.4. OPCIÓN MOSTRAR CUENTAS POR COBRAR.
def mostrarCuentasPorCobrar():
    inputCyanNegrita("")

#--------------------------------------1.2.5. OPCIÓN ANÁLISIS DE CUENTAS POR COBRAR.
def analisisDeCuentasPorCobrar():
    inputCyanNegrita("")

menuPrincipal()
