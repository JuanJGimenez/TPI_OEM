# Bot Cambio Atlántico

Simulador de un chatbot que automatiza la atención de una casa de cambio (compra/venta de USD y EUR). Trabajo Práctico Integrador — Organización Empresarial (TUP a Distancia).

> Simulación académica: el bot informa cotizaciones y registra operaciones simuladas. No opera dinero real ni brinda asesoramiento financiero.

## Funcionalidades

- Informa la cotización (compra/venta) de USD y EUR.
- Permite comprar, vender o solo consultar la cotización.
- Verifica disponibilidad de stock (divisa) o caja (pesos) antes de cerrar.
- Registra cada operación en la planilla y descuenta el stock.
- Si el usuario ingresa datos incorrectos tres veces seguidas, el bot lo deriva a un operador.

## Tecnología

- Lenguaje: Python, usando funciones, diccionarios e if/elif. Sin clases.
- Persistencia: planilla Excel leída y escrita con la librería `openpyxl`.
- Interfaz: simulador por consola, no requiere conexión a internet ni token.

## Archivos

El proyecto contiene los siguientes archivos:

- **bot_cambio.py**: programa principal con el simulador y la máquina de estados.
- **CambioAtlantico_BD.xlsx**: base de datos simulada con cotizaciones, stock y operaciones.
- **README.md**: este archivo con las instrucciones del proyecto.

## Cómo ejecutarlo

1. Instalar la librería necesaria:
```bash
   pip install openpyxl
```
2. Asegurarse de que bot_cambio.py y CambioAtlantico_BD.xlsx estén en la misma carpeta.
3. Ejecutar el programa:
```bash
   python bot_cambio.py
```
4. Escribir `operar` para iniciar y seguir las instrucciones del bot. Escribir `salir` para terminar.

## Estados del bot

El bot maneja cinco estados internos que representan cada paso de la conversación:

`REPOSO → MENU_MONEDA → MENU_OPERACION → ESPERA_MONTO → ESPERA_CONFIRMACION`

Cada estado corresponde a un punto de espera del diagrama BPMN del proyecto.

## Qué hace el bot paso a paso

1. Muestra la cotización del día para USD y EUR.
2. Pregunta si el usuario quiere operar o solo consultar.
3. Si quiere operar, pregunta qué moneda y si desea comprar o vender.
4. Solicita el monto y valida que sea un número mayor a cero.
5. Verifica si hay stock disponible para concretar la operación.
6. Muestra un resumen y pide confirmación antes de registrar.
7. Si el usuario confirma, registra la operación y actualiza el stock.

## Posible extensión a Telegram

La lógica del bot está pensada para poder migrarse a Telegram sin reescribir todo desde cero. Las funciones principales (`leer_cotizaciones`, `hay_disponibilidad`, `calcular_total` y `guardar_operacion`) y la máquina de estados se reutilizarían tal cual, conectándolas a la librería `python-telegram-bot` en lugar de la consola.

## Integrantes

- Emiliano Rojas
- Juan José Giménez
