const {payload} = require('./payload')
let xl = require('excel4node');
let wb = new xl.Workbook();
let options = {
    margins: {
      left: 1.5,
      right: 1.5,
    },
    sheetView: {
        'showGridLines': false, // Flag indicating whether the sheet should have gridlines enabled or disabled during view
        'zoomScale': 85, // Defaults to 100
    }
  };

let border =   {
    top: { style: 'medium', color: "#DDDDDD" }, 
    bottom: { style: 'medium', color: "#DDDDDD" },
    left: { style: 'medium', color: "#DDDDDD" },
    right: { style: 'medium', color: "#DDDDDD" }
} 
let styleHeader = wb.createStyle({
    font: {
        color: '#FFFFFF',
        size: 11,
        name: 'Trebuchet MS'
    },
    fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: '#1788D7',
        bgColor: '#1788D7'
    },
    border,
    alignment: {
        wrapText: true,
        horizontal: 'center',
        vertical: 'center'
    }
});

const styleAditionalInfo = (bold = false) => {
    return wb.createStyle({
        font: {
            color: 'black',
            name: 'Calibri',
            size: 10,
            bold: bold,
        },
    });
}

let styleData = wb.createStyle({
    font: {
        color: 'black',
        size: 11,
        name: 'Trebuchet MS',
    },
    border,
    alignment: {
        wrapText: true,
        vertical: 'center'
    },
});

let ws = wb.addWorksheet('sheetname', options);
let rowPointer = 5

const writingDataOperation = (dataType, column, row, data, style, index) => {
    try{
        verifyType(data, index)
        if(dataType == 'string'){
            return ws.cell(row, column).string(data).style(style)
        }
        if(dataType == 'number'){
            return ws.cell(row, column).number(data).style(style)
        }
    } catch(error){
        console.error(`[ writingDataOperation ]  ${error}`)
    }
}

const verifyType = (data, index) => {
    const contraste = payload.cabeceras[index]
    if(typeof(data) != contraste.tipo){
        throw `El tipo de dato no es el mismo indicado en cabecera. El tipo en cabecera es ${constraste.tipo}, 
        el dato es ${dato}, el tipo del dato errado es: ${typeof(data)}`
    }
    return typeof(data)
}

const writeCell = (column, row, data, style) => {
    ws.cell(column, row)
        .string(data)
        .style(style);
}

const writingAditionalInfoProcess = (payload) => {
    const informacionAdicionalData = Object.entries(payload.informacionAdicional)
    for (const iterator of informacionAdicionalData) { 
        const capitalized = iterator[0].charAt(0).toUpperCase() + iterator[0].slice(1)
        writeCell(rowPointer,1, `${capitalized}:`, styleAditionalInfo(true) )
        writeCell(rowPointer,2, iterator[1], styleAditionalInfo() )
        rowPointer++
    }
}

const writingDataProcess = (payload) => {
    let column = 1
    let indexCounter = 0
    const data = payload.datos
    data.forEach( (value, index) => {
        rowPointer++
        for (const iterator of value) { 
            // writeCell(rowPointer, column, iterator.toString(), styleData )
            writingDataOperation(typeof(iterator), column, rowPointer, iterator, styleData, indexCounter)
            ws.row(rowPointer).setHeight(30)
            column++
            indexCounter++         
        }
        indexCounter = 0
        column=1;
    })
}

const writingHeaderProcess = (payload) => {
    const { cabeceras } = payload
    const cabecerasData = Object.entries(cabeceras)
    rowPointer++
    for (const iterator of cabecerasData) {
        const column = parseInt(iterator[0]) + 1
        ws.column(column).setWidth(20)
        ws.row(rowPointer).setHeight(30)
        writeCell(rowPointer, column, iterator[1].nombre, styleHeader )
        console.log(iterator) 
    }
    
}

ws.addImage({
    path: './logoSura.png',
    type: 'picture',
    position: {
      type: 'oneCellAnchor',
      from: {
        col: 1,
        colOff: '0.1in',
        row: 1,
        rowOff: '0.2in',
      },
    },
});

const procesarExcel = (payload) => {
    try{

        writingAditionalInfoProcess(payload)
        writingHeaderProcess(payload)
        writingDataProcess(payload)
        wb.write('ExcelFile.xlsx');
        return wb.writeToBuffer()
    }catch(e) {
        console.error(`Hubo un error ${e}`)
    }
}

/* writingAditionalInfoProcess()
writingHeaderProcess()
writingDataProcess() */

module.exports = {
    procesarExcel
}


