# -*- coding: utf-8 -*-
# Simulador del bot de Cambio Atlántico (casa de cambio USD / EUR).
# Trabajo Práctico Integrador - Organización Empresarial.
#
# Usa solo conceptos basicos: variables, diccionarios, funciones, if/elif y while.
# El "estado" es un simple texto ("REPOSO", "MENU_MONEDA", etc.) que le da memoria
# al bot, igual que en el diagrama de la maquina de estados.
#
# Requisitos:  pip install openpyxl
# Ejecutar:    python bot_cambio.py

from openpyxl import load_workbook
from datetime import datetime

ARCHIVO = "CambioAtlantico_BD.xlsx"  # nuestra base de datos simulada
MAX_INTENTOS = 3
OPERADOR = "+54 9 11 0000-0000"


# ----------------------- Funciones que leen/escriben la planilla -----------------------


def leer_cotizaciones():
    """
    Devuelve un diccionario: {'USD': {'compra': 1180, 'venta': 1230}, ...}
    """
    libro = load_workbook(ARCHIVO, data_only=True)
    hoja = libro["Cotizaciones"]
    cotizaciones = {}
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        moneda = fila[0]
        if moneda:
            cotizaciones[moneda] = {"compra": fila[1], "venta": fila[2]}
    return cotizaciones


def leer_stock():
    """
    Devuelve un diccionario: {'USD': 25000, 'EUR': 12000, 'ARS': 40000000}
    """
    libro = load_workbook(ARCHIVO, data_only=True)
    hoja = libro["Stock"]
    stock = {}
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        moneda = fila[0]
        if moneda:
            stock[moneda] = fila[1]
    return stock


def hay_disponibilidad(moneda, operacion, monto):
    """
    Compra: hace falta stock de la divisa. Venta: hace falta caja en pesos.
    Devuelve True si hay disponibilidad, False si no.
    """
    stock = leer_stock()
    cotizaciones = leer_cotizaciones()
    if operacion == "comprar":
        return stock[moneda] >= monto
    else:  # vender: la casa entrega pesos
        pesos_necesarios = monto * cotizaciones[moneda]["compra"]
        return stock["ARS"] >= pesos_necesarios


def calcular_total(moneda, operacion, monto):
    """
    Devuelve (total_en_pesos, cotizacion_aplicada).
    Para comprar, se aplica la cotizacion de venta. Para vender, la cotizacion de compra.
    """

    cotizaciones = leer_cotizaciones()
    if operacion == "comprar":
        tasa = cotizaciones[moneda]["venta"]
    else:
        tasa = cotizaciones[moneda]["compra"]
    return monto * tasa, tasa


def guardar_operacion(operacion, moneda, monto):
    """
    Agrega la operacion a la planilla y ajusta el stock. Devuelve (total, cotizacion_aplicada).
    """
    total, tasa = calcular_total(moneda, operacion, monto)
    libro = load_workbook(ARCHIVO)
    hoja_op = libro["Operaciones"]
    nuevo_id = hoja_op.max_row  # la fila 1 es el encabezado
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    hoja_op.append(
        [
            nuevo_id,
            fecha,
            "cliente",
            operacion,
            moneda,
            monto,
            tasa,
            total,
            "COMPLETADA",
        ]
    )
    # Actualizar el stock segun la operacion
    hoja_st = libro["Stock"]
    for fila in range(2, hoja_st.max_row + 1):
        nombre = hoja_st.cell(fila, 1).value
        if operacion == "comprar":
            if nombre == moneda:
                hoja_st.cell(fila, 2).value -= monto  # entrega divisa
            if nombre == "ARS":
                hoja_st.cell(fila, 2).value += total  # recibe pesos
        else:
            if nombre == moneda:
                hoja_st.cell(fila, 2).value += monto  # recibe divisa
            if nombre == "ARS":
                hoja_st.cell(fila, 2).value -= total  # entrega pesos
    libro.save(ARCHIVO)
    return total, tasa


# ----------------------------- Programa principal -----------------------------

print("=== Bot Cambio Atlántico ===")
print("Escribí 'operar' para comenzar, o 'salir' para terminar.\n")

estado = "REPOSO"  # estado actual (la memoria del bot)
moneda = ""  # datos que vamos juntando durante la conversacion
operacion = ""
monto = 0
intentos = 0  # cuenta los errores seguidos para derivar a un operador

while True:
    texto = input("Vos: ").strip()
    cliente = texto.lower()

    if cliente == "salir":
        print("Bot: ¡Hasta luego!")
        break

    # operar reinicia la conversacion desde cualquier estado
    if cliente == "operar":
        estado = "MENU_MONEDA"
        moneda = ""
        operacion = ""
        monto = 0
        intentos = 0
        print("Bot: Bienvenido a Cambio Atlántico. ¿Con qué moneda desea operar? (USD o EUR)")
        continue

    if estado == "REPOSO":
        print("Bot: Escribí 'operar' para comenzar.")

    elif estado == "MENU_MONEDA":
        if cliente.upper() in ("USD", "EUR"):
            moneda = cliente.upper()
            intentos = 0
            estado = "MENU_OPERACION"
            cot = leer_cotizaciones()[moneda]
            print(
                "Bot: Cotización "
                + moneda
                + ": compra $"
                + str(cot["compra"])
                + " / venta $"
                + str(cot["venta"])
            )
            print("Bot: ¿Querés comprar / vender o solo consultar?")
        else:
            intentos += 1
            if intentos >= MAX_INTENTOS:
                print(
                    "Bot: Te derivo con un operador: " + OPERADOR + ". Escribí 'operar' para volver a intentarlo."
                )
                estado = "REPOSO"
                intentos = 0
            else:
                print("Bot: Por ahora solo operamos USD y EUR. Elegí una de las dos.")

    elif estado == "MENU_OPERACION":
        if cliente in ("comprar", "vender"):
            operacion = cliente
            intentos = 0
            estado = "ESPERA_MONTO"
            print("Bot: ¿Qué monto de " + moneda + " querés " + operacion + "?")
        elif cliente in ("consultar", "no", "salir"):
            estado = "REPOSO"
            intentos = 0
            print(
                "Bot: ¡Listo! Esa es la cotización de hoy. Escribí 'operar' cuando quieras comenzar."
            )
        else:
            intentos += 1
            if intentos >= MAX_INTENTOS:
                print(
                    "Bot: Te derivo con un operador: " + OPERADOR + ". Escribí 'operar' para volver a intentarlo."
                )
                estado = "REPOSO"
                intentos = 0
            else:
                print("Bot: Elegí: comprar, vender o consultar.")

    elif estado == "ESPERA_MONTO":
        if not cliente.isdigit():
            print("Bot: Ingresá el monto en números, por ejemplo: 100.")
        elif int(cliente) <= 0:
            print("Bot: El monto debe ser mayor a cero.")
        elif not hay_disponibilidad(moneda, operacion, int(cliente)):
            estado = "MENU_OPERACION"
            print(
                "Bot: No tenemos esa cantidad disponible. Elegí otra operación o usá 'operar'."
            )
        else:
            monto = int(cliente)
            intentos = 0
            total, tasa = calcular_total(moneda, operacion, monto)
            estado = "ESPERA_CONFIRMACION"
            print(
                "Bot: "
                + operacion.capitalize()
                + " de "
                + str(monto)
                + " "
                + moneda
                + " a $"
                + str(tasa)
                + ". Total: $"
                + str(total)
                + ". ¿Confirmás? (sí/no)"
            )

    elif estado == "ESPERA_CONFIRMACION":
        if cliente in ("si", "sí", "s"):
            total, tasa = guardar_operacion(operacion, moneda, monto)
            estado = "REPOSO"
            print("Bot: Operación registrada. Total: $" + str(total) + ".")
            print(
                "Bot: Gracias por operar con Cambio Atlántico. Escribí 'operar' para otra operación."
            )
        elif cliente in ("no", "n", "cancelar"):
            estado = "REPOSO"
            print("Bot: Operación cancelada. Escribí 'operar' cuando quieras.")
        else:
            print("Bot: Respondé sí o no para confirmar.")
