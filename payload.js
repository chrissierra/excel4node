

const payload = {
    "informacionAdicional": {
        "nombre": "PABLO MATTE",
        "rut": "138298086",
        "operacion": "Inversión Extranjera/Descargar detalle excel"
    },
    "datos": [
        [
            "Money Market",
            "-",
            101800.73,
            1,
            713,
            101800.73,
            72537092
        ],
        [
            "Money Market2",
            "-",
            201800.73,
            2,
            213,
            2101800.73,
            22537092
        ],
        [
            "Soney Market2",
            "-",
            301800.73,
            3,
            313,
            3101800.73,
            32537092
        ]
    ],
    "cabeceras": [
        {
            "tipo": "string",
            "nombre": "Instrumentos"
        },
        {
            "tipo": "string",
            "nombre": "Ticker"
        },
        {
            "tipo": "number",
            "nombre": "Cantidad"
        },
        {
            "tipo": "number",
            "nombre": "Precio en dólares"
        },
        {
            "tipo": "number",
            "nombre": "Precio en pesos"
        },
        {
            "tipo": "number",
            "nombre": "Saldo en dólares"
        },
        {
            "tipo": "number",
            "nombre": "Saldo en pesos"
        }
    ]
}

module.exports = { payload }