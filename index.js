const XLSX = require('xlsx-js-style')

const userInfo = [
    ["Nombre Cliente","DANIELA ALEJANDRA CANDIA ZAMORANO"],
    ["Rut Cliente","160987933"],
    ["Fecha"," 28 / 11 / 2022 3:36"],
    ["Operación", "Nombre transación" ],
]

const cabeceras = [
    {tipo: 'string', nombre: 'Tipo Transacción'},
    {tipo: 'string', nombre: 'Número de cuota'},
    {tipo: 'number', nombre: 'Valor cuota'},
    {tipo: 'number', nombre: 'Saldo'},
    {tipo: 'string', nombre: 'Nombre fondo'},
    {tipo: 'number', nombre: 'N° Solicitud'},
    {tipo: 'string', nombre: 'Canal de ingreso'},
    {tipo: 'date', nombre: 'Fecha Ingreso'},
    {tipo: 'string', nombre: 'Estado Actual'},
    {tipo: 'date', nombre: 'Fecha Cierre'},
]


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
)

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

const setAddresA1B1TypeToHeadersObject = (cabecera) => {
    // // console.log(`setAddresA1B1TypeToHeadersObject`,cabecera)
    cabecera.forEach( (element, index) => {
        element.a1b1 = index
    });
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
    // const worksheet2 = hideGridLines()
    const workSheet = XLSX.utils.aoa_to_sheet(userInfo, {origin: 3})
    setAddresA1B1TypeToHeadersObject(cabeceras)
    // const workSheet2 = XLSX.utils.json_to_sheet(dto);
    const workBook = XLSX.utils.book_new();
    for (const property in workSheet) {
        //// // console.log('Primera pasada', workSheet[property] )
        workSheet[property].s = { 
            font: { name: "Trebuchet MS", sz: 11 }, 
            fill: { fgColor: { rgb: "FFFFFF" } },
            border: { style: 'thin', color: "FFFFFF" }
        }
        if(property.includes('1') && property.length == 2){          
            workSheet[property].s = headersStyles
        }
    }
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
                // // console.log(workSheet[property])
            }
            if(workSheet[property].z){
                workSheet[property].t = 'd'
                // console.log(workSheet[property])
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
            workSheet[property].s = headersStyles
        }
        // // console.log(`Segunda ${count}`, {hoja: workSheet[property], property} )
    }


    // console.log(`Metadata: `, {metaData, workSheet, cols: workSheet['!cols']})
    
    XLSX.utils.book_append_sheet(workBook, workSheet, "Transacciones APV Fondos Mutuos")
    // Generate buffer
    XLSX.write(workBook, { bookType: 'xlsx', type: "buffer" })
    // Binary string
    XLSX.write(workBook, { bookType: "xlsx", type: "binary" })

    XLSX.writeFile(workBook, "Tabla.xlsx")

}
convertJsonToExcel()