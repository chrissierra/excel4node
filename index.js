const XLSX = require('xlsx-js-style')

/* const userInfo = [
    ["Nombre Cliente","DANIELA ALEJANDRA CANDIA ZAMORANO"],
    ["Rut Cliente","160987933"],
    ["Fecha"," 28 / 11 / 2022 3:36"],
    ["Operación", "Nombre transación" ],
] */
/* 
const dto = 
new Array(3).fill(
    {
        "Tipo Transacción": 'Trx Tipo 1',
        "Número de cuota": 5,
        "Valor cuota": 100000,
        "Saldo": '$10000',
        "Nombre fondo": 'Fondo A',
        "N° Solicitud": '21341234',
        "Canal de ingreso": 'Canal WEB',
        "Fecha Ingreso": new Date().toLocaleDateString(),
        "Estado Actual": 'aprobado',
        "Fecha Cierre":  new Date().toLocaleDateString(),
    }
) */

const buildDTO = (payload) => {
    let datos = payload.datos
    let cabeceras = payload.cabeceras
    let objectResult = {};
    let result = []
    datos.map( (values) => {
        cabeceras.forEach((key, i) => objectResult[key.nombre] = values[i]);
        result.push({...objectResult})
    } )
    return result
}


const buildAditionalInfo = (payload) => {
    return Object.entries(payload.informacionAdicional)
}


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

const hideGridLines = () => {
    let columnsDummy = new Array(30).fill('dummyData')
    columnsDummy = columnsDummy.map((value, index)=> `${index}${value}`)
    let rowsDummys = []
    rowsDummys.push(columnsDummy)
    for (let index = 0; index < 5000; index++) {
        rowsDummys.push(columnsDummy)
    }
    
    let workSheet= XLSX.utils.aoa_to_sheet(rowsDummys, {origin: 3})
    for (const property in workSheet) {
        workSheet[property].s = { 
            fill: { fgColor: { rgb: "FFFFFF" } },
            border: { style: 'thin', color: "FFFFFF" }
        }
    }
    console.log(workSheet)
    for (const property in workSheet) {
        workSheet[property].v = ''
    }
    return workSheet;
}


const convertJsonData = (jsonData) => {
    let result = []
    result.push([])
    result.push(Object.keys(jsonData[0]))
    for (let name of jsonData) {
        result.push(Object.values(name))
    }
    return result
}



const getDtoMetaData = (dto) => {
    const DTOMetaData =  {
        lengthRow: Object.keys(dto[0]).length,
        lastColumnValue() { return Object.keys(dto[0])[this.lengthRow-1]},
        initTable: Object.keys(dto[0])[0],
        getRowFromAddress(address) { return address.match(/\d/g).join("")},
        getInitTableA1B1(address, range) { 
            let response = {}
            const cellValue = range[address].v
            if(cellValue == this.initTable){
                const initRowNumber = parseInt(this.getRowFromAddress(address))
                response = {
                    initRowAddres: address,
                    initRowNumber,
                    endRowNumber: initRowNumber + dto.length,
                    endRowAddress() { return `${this.addresLastHeaderColumn.replace(/[^a-zA-Z]+/g, '')}${this.endRowNumber}` }, 
                    lastColumnValue: this.lastColumnValue(),
                    lengthRow: this.lengthRow,
                    addresLastHeaderColumn: Object.entries(range).find(([key, value]) => value.v == this.lastColumnValue())[0]
                } 
                return response
            }
            return false;
        },
        getColumnEndTable(address, range) {},
        endTable: Object.values(dto[dto.length-1])[Object.keys(dto[0]).length-1]
    }

    return DTOMetaData;
}


const borderStyle = {
    top: { style: 'medium', color: { rgb: "DDDDDD" } }, 
    bottom: { style: 'medium', color: { rgb: "DDDDDD" } },
    left: { style: 'medium', color: { rgb: "DDDDDD" } },
    right: { style: 'medium', color: { rgb: "DDDDDD" } }
}
const headersStyles = { font: { name: "Trebuchet MS", sz: 11, color: {rgb: 'FFFFFF'} } , fill: { fgColor: { rgb: "1788D7" } },
border: borderStyle }
const tableStyles = { 
    font: { name: "Trebuchet MS", sz: 11 }, 
    fill: { fgColor: { rgb: "FFFFFF" } },
    border: borderStyle
}

const convertJsonToExcel = () => {

    const userInfo = buildAditionalInfo(payload)
    let dto = buildDTO(payload)
    console.log(buildDTO(payload))
    
    // const workSheet = XLSX.utils.aoa_to_sheet(userInfo, {origin: 3})
    // const workBook = XLSX.utils.book_new();
    
    const workBook = XLSX.readFile('Tabla.xlsx', {cellStyles: true, type: 'file', bookFiles: true})
    const workSheet = workBook.Sheets['Transacciones APV Fondos Mutuos']
    const workSheet2 = workBook.Sheets['Transacciones APV Fondos Mutuos']
    // console.log(workBook)
    console.log(workBook.Themes.raw)
    // XLSX.utils.sheet_add_aoa(workSheet, convertJsonData(dto), { origin: -1 });
    XLSX.utils.sheet_add_aoa(workSheet, userInfo, { origin: 6 });
    XLSX.utils.sheet_add_aoa(workSheet, convertJsonData(dto), { origin: -1 });
    let metaData = {}
    let insideTable = false;
    let insideHeaders = false
    for (const property in workSheet) {
        if(getDtoMetaData(dto).getInitTableA1B1(property, workSheet)) {
            metaData = {... getDtoMetaData(dto).getInitTableA1B1(property, workSheet) }
        }
        if(insideTable){
            workSheet[property].s = tableStyles
           
            if(workSheet[property].t=='n'){
                workSheet[property].t = 'n'
            }
            if(workSheet[property].z){
                workSheet[property].t = 'd'
            }
        }
        if(insideHeaders || getDtoMetaData(dto).getInitTableA1B1(property, workSheet)){
            // Apply headers styles
            workSheet[property].s = headersStyles
            insideHeaders = true
        }
        if(metaData.addresLastHeaderColumn == property){
            insideHeaders = false
            insideTable = true
            workSheet[property].h = workBook.Themes.raw
        }
    }

    
    
    XLSX.utils.book_append_sheet(workBook, workSheet2, "aver")
    // Generate buffer
    XLSX.write(workBook, { bookType: 'xlsx', type: "buffer", cellStyles: true })
    // Binary string
    XLSX.write(workBook, { bookType: "xlsx", type: "binary" , cellStyles: true})
    XLSX.writeFile(workBook, "Tabla2.xlsx", {cellStyles: true})
}
convertJsonToExcel()