# Bot Cambio Atlántico

Simulador de un chatbot que automatiza la atención de una casa de cambio
(compra/venta de USD y EUR). Trabajo Práctico Integrador — Organización
Empresarial (TUP a Distancia).

> Simulación académica: el bot informa cotizaciones y registra operaciones
> simuladas. No opera dinero real ni brinda asesoramiento financiero.

## Funcionalidades
- Informa la cotización (compra/venta) de USD y EUR.
- Permite comprar, vender o solo consultar la cotización.
- Verifica disponibilidad de stock (divisa) o caja (pesos) antes de cerrar.
- Registra cada operación en la planilla y descuenta el stock.
- Maneja errores de entrada mediante una máquina de estados, y deriva a un
  operador tras 3 intentos inválidos.

## Tecnología
- Lenguaje: Python (funciones, diccionarios, if/elif y while; sin clases).
- Persistencia: planilla Excel leída con la librería `openpyxl`.
- Interfaz: simulador por consola (no requiere conexión ni token).

## Archivos
```
.
├── bot_cambio.py             # Programa principal (simulador + máquina de estados)
├── CambioAtlantico_BD.xlsx   # Base de datos simulada (cotizaciones, stock, operaciones)
└── README.md
```

## Cómo ejecutarlo
1. Instalar la librería para leer la planilla:
   ```bash
   pip install openpyxl
   ```
2. Ejecutar el simulador (el archivo .xlsx debe estar en la misma carpeta):
   ```bash
   python bot_cambio.py
   ```
3. Escribir `/start` para comenzar y seguir las indicaciones. `salir` para terminar.

## Estados del bot
`REPOSO → MENU_MONEDA → MENU_OPERACION → ESPERA_MONTO → ESPERA_CONFIRMACION`

Cada estado se corresponde con un punto de espera del diagrama BPMN.
Ver la matriz de transiciones completa en el informe del proyecto.

## Posible extensión a Telegram
El programa se puede llevar a Telegram reutilizando las mismas funciones
(`leer_cotizaciones`, `hay_disponibilidad`, `calcular_total`, `guardar_operacion`)
y la misma lógica de estados, conectándolas a la librería `python-telegram-bot`.

## Integrantes
- [ Nombre y Apellido 1 ]
- [ Nombre y Apellido 2 ]
