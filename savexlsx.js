/*var excel = require('excel4node');

// Create a new instance of a Workbook class
var workbook = new excel.Workbook();

// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet 1');
var worksheet2 = workbook.addWorksheet('Sheet 2');

// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

// Set value of cell A1 to 100 as a number type styled with paramaters of style
worksheet.cell(1,1).number(100).style(style);

// Set value of cell B1 to 300 as a number type styled with paramaters of style
worksheet.cell(1,2).number(200).style(style);

// Set value of cell C1 to a formula styled with paramaters of style
worksheet.cell(1,3).formula('A1 + B1').style(style);

// Set value of cell A2 to 'string' styled with paramaters of style
worksheet.cell(-2,1).string('string').style(style);

// Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
worksheet.cell(3,1).bool(true).style(style).style({font: {size: 14}});
worksheet.cell(4,1).string("hiii");
workbook.write('Excel.xlsx');*/
const express = require('express');
const bodyParser = require('body-parser')

const app = express();

app.use(bodyParser.urlencoded({extended: true}));

var temp = req.body.email; console.log(temp)


var Workbook =require ('excel4node');
const wb = new Workbook();
const ws = wb.addWorksheet('Worksheet Name');
const data = [
 {
    
    "email":email.value,
    "mobile":"1234567890"
 }
]
const headingColumnNames = [
 
    "Email",
    "Mobile",
]
//Write Column Title in Excel file
let headingColumnIndex = 1;
headingColumnNames.forEach(heading => {
    ws.cell(1, headingColumnIndex++)
        .string(heading)
});
//Write Data in Excel file
let rowIndex = 2;
data.forEach( record => {
    let columnIndex = 1;
    Object.keys(record ).forEach(columnName =>{
        ws.cell(rowIndex,columnIndex++)
            .string(record [columnName])
    });
    rowIndex++;
}); 
wb.write('data.xlsx');