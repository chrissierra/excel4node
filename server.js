const express = require('express')
const app = express()
const {procesarExcel} = require('./excel4node')
var bodyParser = require('body-parser')
// create application/json parser
var jsonParser = bodyParser.json()
 
// create application/x-www-form-urlencoded parser
// var urlencodedParser = bodyParser.urlencoded({ extended: false })


// POST /login gets urlencoded bodies
app.post('/', jsonParser, function (req, res) {
    console.log(procesarExcel)
    const buffer = procesarExcel(req.body)
    buffer.then( b => res.send(b.toString('base64')))
  })

  app.listen(8989, () => {
    console.log(`Example app listening on port ${8989}`)
  })