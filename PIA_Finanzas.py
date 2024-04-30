#Módulos.
import pandas as pd
import sqlite3
from sqlite3 import Error
import sys

#Lee el excel.
try:
    df = pd.read_excel('Antiguedad_de_saldos.xlsx', engine = 'openpyxl')
    print("Lectura del excel correctamente.")
except Error as e:
        print(f'Se produjo el siguiente error: {e}')
except Exception:
        print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

#Copia el df a otro sin nombres duplicados.
try:
    df_copia1 = df[['Nombre']].copy()
    df_copia1 = df_copia1.drop_duplicates().reset_index(drop = True)
except Error as e:
        print(f'Se produjo el siguiente error: {e}')
except Exception:
        print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

#Copia el df a otro sin nombres, para cuentas por cobrar.
try:
    df_copia2 = df.copy()
    df_copia2 = df_copia2.drop_duplicates().reset_index(drop = True)
except Error as e:
        print(f'Se produjo el siguiente error: {e}')
except Exception:
        print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')

#Creación o cargado de tablas.
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
            print ('Creación o cargado de tablas realizado correctamente.')                         
    except Error as e:
        print(f'Se produjo el siguiente error: {e}')
    except Exception:
        print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
    finally:
        conn.close()
creacion_tablas()

#Insertar datos en CLIENTES solo si no existen.
try:
    with sqlite3.connect('CuentasPorCobrar.db') as conn:
        mi_cursor = conn.cursor()
        for nombre_cliente in df_copia1['Nombre']:
            mi_cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE CLAVE_CLIENTE=?", (nombre_cliente,))
            count = mi_cursor.fetchone()[0]
            if count == 0:
                mi_cursor.execute("INSERT INTO CLIENTES(NOMBRECLIENTE) VALUES (?)", (nombre_cliente,))
        print('Datos insertados en la tabla CLIENTES correctamente.')
except Error as e:
    print(f'Se produjo el siguiente error: {e}')
except Exception:
    print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
finally:                                    
    conn.close()

# Insertar datos en CuentasPorCobrar solo si CLAVE_CLIENTE existe en CLIENTES y CLAVE_GUIA no existe.
try:
    with sqlite3.connect('CuentasPorCobrar.db') as conn:
        mi_cursor = conn.cursor()
        for index, row in df.iterrows():
            mi_cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE CLAVE_CLIENTE=?", (row['Cliente'],))
            existencia2 = mi_cursor.fetchone()[0]

            mi_cursor.execute("SELECT COUNT(*) FROM CuentasPorCobrar WHERE CLAVE_GUIA=?", (row['Guía'],))
            count_guia = mi_cursor.fetchone()[0]

            if existencia2 > 0 and count_guia == 0:
                fecha = row['Fecha'].strftime('%Y-%m-%d')  # Convierte la fecha a una cadena de texto.
                mi_cursor.execute("INSERT INTO CuentasPorCobrar(CLAVE_GUIA, DIAS, FECHA, TOTAL, CLAVE_CLIENTE) VALUES (?, ?, ?, ?, ?)", 
                                  (row['Guía'], row['Días '], fecha, row['Total'], row['Cliente']))
        print('Datos insertados en la tabla CuentasPorCobrar correctamente.')
except Error as e:
    print(f'Se produjo el siguiente error: {e}')
except Exception:
    print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
finally:                                    
    conn.close()

'''#Ver registros.
try:
    with sqlite3.connect('CuentasPorCobrar.db') as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("SELECT * FROM CLIENTES")
        registros = mi_cursor.fetchall()
        for registro in registros:
            print(registro)
except Error as e:
    print(f'Se produjo el siguiente error: {e}')
except Exception:
    print(f'Se produjo el siguiente error: {sys.exc_info()[0]}')
finally:                                    
    conn.close()'''