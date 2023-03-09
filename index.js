const XLSX = require('xlsx-js-style')

const userInfo = [
    ["Nombre Cliente","DANIELA ALEJANDRA CANDIA ZAMORANO"],
    ["Rut Cliente","160987933"],
    ["Fecha","28/11/2022 3:36"],
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
        "Fecha Ingreso": new Date(),
        "Estado Actual": 'aprobado',
        "Fecha Cierre": '15/12/2022',
    }
)

const hideGridLines = (workSheet) => {
    for (const property in workSheet) {
        workSheet[property].s = { 
            fill: { fgColor: { rgb: "FFFFFF" } },
            border: { style: 'thin', color: "FFFFFF" }
        }
    }
}

const setAddresA1B1TypeToHeadersObject = (cabecera) => {
    // console.log(`setAddresA1B1TypeToHeadersObject`,cabecera)
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
                    endRowAddress() { return `${this.addressLastColumn.replace(/[^a-zA-Z]+/g, '')}${this.endRowNumber}` }, 
                    lastColumnValue: this.lastColumnValue(),
                    lengthRow: this.lengthRow,
                    addressLastColumn: Object.entries(range).find(([key, value]) => value.v == this.lastColumnValue())[0]
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

const convertJsonToExcel = () => {
    const workSheet = XLSX.utils.aoa_to_sheet(userInfo, {origin: 3})
    setAddresA1B1TypeToHeadersObject(cabeceras)
    const workSheet2 = XLSX.utils.json_to_sheet(dto);
    const workBook = XLSX.utils.book_new();
    for (const property in workSheet) {
        //// console.log('Primera pasada', workSheet[property] )
        workSheet[property].s = { 
            font: { name: "Trebuchet MS", sz: 11 }, 
            fill: { fgColor: { rgb: "FFFFFF" } },
            border: { style: 'thin', color: "FFFFFF" }
        }
        if(property.includes('1') && property.length == 2){          
            workSheet[property].s = { font: { name: "Trebuchet MS", sz: 11, color: {rgb: 'FFFFFF'} } , fill: { fgColor: { rgb: "1788D7" } } }
        }
    }
    XLSX.utils.sheet_add_aoa(workSheet, convertJsonData(dto), { origin: -1 });
    let count = 0
    let metaData = {}
    for (const property in workSheet) {
        count++;
        if(getDtoMetaData(dto).getInitTableA1B1(property, workSheet)){
            //console.log(`Addres primer valor tabla: (${property})`, getDtoMetaData(dto).getInitTableA1B1(property, workSheet))
            //console.log(getDtoMetaData(dto).getInitTableA1B1(property, workSheet).endRowAddress())
            metaData = {... getDtoMetaData(dto).getInitTableA1B1(property, workSheet) }
        }
        // console.log(`Segunda ${count}`, {hoja: workSheet[property], property} )
    }


    console.log(`Metadata: `, metaData)
    
    XLSX.utils.book_append_sheet(workBook, workSheet, "students")
    // Generate buffer
    XLSX.write(workBook, { bookType: 'xlsx', type: "buffer" })
    // Binary string
    XLSX.write(workBook, { bookType: "xlsx", type: "binary" })

    XLSX.writeFile(workBook, "studentsData.xlsx")

}
convertJsonToExcel()