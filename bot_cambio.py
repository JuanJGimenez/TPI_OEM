# Simulador del bot de Cambio Atlántico (casa de cambio USD / EUR).
# Trabajo Práctico Integrador - Organización Empresarial.
#
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
    Lee las cotizaciones actuales de las divisas desde la base de datos Excel.

    Accede a la hoja 'Cotizaciones' del archivo Excel configurado y extrae
    los valores de compra y venta para cada moneda disponible.

    Returns:
        dict: Un diccionario con las monedas como llaves y otro diccionario
            con sus valores de 'compra' y 'venta'.
            Ejemplo: {'USD': {'compra': 1180, 'venta': 1230}}
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
    Lee los fondos disponibles de cada moneda en la base de datos Excel.

    Accede a la hoja 'Stock' para obtener las existencias físicas de las
    divisas extranjeras y de la caja en pesos argentinos (ARS).

    Returns:
        dict: Un diccionario con el código de la moneda y su cantidad disponible.
            Ejemplo: {'USD': 25000, 'EUR': 12000, 'ARS': 40000000}
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
    Verifica si la casa de cambio tiene fondos suficientes para la transacción.

    Si es una compra del cliente, valida si hay suficiente stock de esa divisa.
    Si es una venta del cliente, calcula el total en pesos y valida si la casa
    tiene suficiente efectivo en ARS para pagarle.

    Args:
        moneda (str): El código de la divisa elegida (ej. 'USD', 'EUR').
        operacion (str): El tipo de transacción ('comprar' o 'vender').
        monto (int): La cantidad de divisa extranjera que se quiere operar.

    Returns:
        bool: True si la casa de cambio puede cubrir la operación, False si no.
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
    Calcula el costo total en pesos y determina la tasa de cambio a aplicar.

    Aplica la cotización de 'venta' si el cliente compra divisas,
    o la cotización de 'compra' si el cliente vende sus divisas.

    Args:
        moneda (str): El código de la divisa elegida (ej. 'USD', 'EUR').
        operacion (str): El tipo de transacción ('comprar' o 'vender').
        monto (int): La cantidad de divisa extranjera a operar.

    Returns:
        tuple: Un par que contiene:
            - total_en_pesos (float): El resultado financiero de la operación.
            - tasa_aplicada (float): El precio unitario por el cual se multiplicó.
    """

    cotizaciones = leer_cotizaciones()
    if operacion == "comprar":
        tasa = cotizaciones[moneda]["venta"]
    else:
        tasa = cotizaciones[moneda]["compra"]
    return monto * tasa, tasa


def guardar_operacion(operacion, moneda, monto):
    """
    Registra la transacción en el historial y actualiza el stock en el Excel.

    Calcula los totales, añade una fila a la hoja 'Operaciones' con la fecha
    y hora actual, e incrementa o reduce los saldos correspondientes en la
    hoja 'Stock' antes de guardar el archivo.

    Args:
        operacion (str): El tipo de transacción ('comprar' o 'vender').
        moneda (str): El código de la divisa elegida (ej. 'USD', 'EUR').
        monto (int): La cantidad de divisa extranjera operada.

    Returns:
        tuple: Un par que contiene:
            - total (float): El monto total final en pesos de la transacción.
            - tasa (float): La cotización final que fue aplicada a la divisa.
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

    # Intentamos guardar el archivo de forma segura. Si el archivo está abierto,
    # atrapamos el error y avisamos al usuario.
    try:
        libro.save(ARCHIVO)
        return total, tasa
    except PermissionError:
        print(
            "\n[ERROR CRÍTICO] No se pudo guardar la operación porque el archivo Excel está abierto."
        )
        print(
            "[Sugerencia] Por favor, cerrá la planilla 'CambioAtlantico_BD.xlsx' y volvé a intentar.\n"
        )
        # Devolvemos None para que el programa principal sepa que la transacción falló
        return None, None


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
        print(
            "Bot: Bienvenido a Cambio Atlántico. ¿Con qué moneda desea operar? (USD o EUR)"
        )
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
                    "Bot: Te derivo con un operador: "
                    + OPERADOR
                    + ". Escribí 'operar' para volver a intentarlo."
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
                    "Bot: Te derivo con un operador: "
                    + OPERADOR
                    + ". Escribí 'operar' para volver a intentarlo."
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
