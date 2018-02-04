 // Require library
 var xl = require('excel4node');

 // Create a new instance of a Workbook class
 var wb = new xl.Workbook();
 /*
 // Add Worksheets to the workbook
 var ws = wb.addWorksheet('Лист 1');
 var ws2 = wb.addWorksheet('Sheet 2');

 // Create a reusable style
 var style = wb.createStyle({
     font: {
         color: '#FF0800',
         size: 12
     },
     numberFormat: '$#,##0.00; ($#,##0.00); -'
 });

 // Set value of cell A1 to 100 as a number type styled with paramaters of style
 ws.cell(1, 1).number(100).style(style);

 // Set value of cell B1 to 300 as a number type styled with paramaters of style
 ws.cell(1, 2).number(200).style(style);

 // Set value of cell C1 to a formula styled with paramaters of style
 ws.cell(1, 3).formula('A1 + B1').style(style);

 // Set value of cell A2 to 'string' styled with paramaters of style
 ws.cell(2, 1).string('Строка на русском').style(style);

 // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
 ws.cell(3, 1).bool(true).style(style).style({ font: { size: 14 } });

 var complexString = [
     'Workbook default font String\n',
     {
         bold: true,
         underline: true,
         italic: true,
         color: 'FF0000',
         size: 18,
         name: 'Courier',
         value: 'Hello'
     },
     ' World!',
     {
         color: '000000',
         underline: false,
         name: 'Arial',
         vertAlign: 'subscript'
     },
     ' All',
     ' these',
     ' strings',
     ' are',
     ' black subsript,',
     {
         color: '0000FF',
         value: '\nbut',
         vertAlign: 'baseline'
     },
     ' now are blue'
 ];
 ws.cell(4, 1).string(complexString);
 ws.cell(5, 1).string('another simple string').style({ font: { name: 'Helvetica' } });
 ws.cell(6, 1).date(new Date()).style({ numberFormat: 'yyyy-mm-dd' });


 wb.write('Excel.xlsx'); */

 var express = require('express');
 var app = express();
 app.get('/', function(req, res) {
     wb.write('Excel.xlsx', res);
 });
 app.listen(3000, function() {
     console.log('Example app listening on port 3000!');
 });