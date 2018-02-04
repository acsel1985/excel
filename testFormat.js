var fs = require('fs');
var _ = require('underscore');
var ExcelJS = require('exceljs');

var Workbook = ExcelJS.Workbook;
var filename = 'testFormat.xlsx';
var wb = new Workbook();
var ws = wb.addWorksheet('blort');

//data-in "Промежуточный формат"
const dataJSON = `{
    "5": {
        "7": {
            "value": "test",
            "fill": { "type": "pattern", "pattern": "solid", "fgColor": { "argb": "FF00FF00" } }
        },
        "2": {
            "value": "font",
            "font": { "name": "Comic Sans MS", "family": 4, "size": 16, "underline": "double", "bold": true }
        }
    }
}`;

//TODO сделать проверку на ошибки JSON
let dataObject = JSON.parse(dataJSON);

_.each(dataObject, function(row, rowIndex) {
    _.each(row, function(cell, cellIndex) {
        console.log(`Строка ${rowIndex}, столбец ${cellIndex}`);
        //TODO сделать проверку rowIndex и cellIndex числа
        let cellIns = ws.getCell(rowIndex, cellIndex);
        /*for (key in cell) {
            cellIns[key] = cell[key];
        }*/
        _.each(cell, function(val, valKey) {
            cellIns[valKey] = val[valKey];
        })
    });
});

wb.xlsx.writeFile(filename)
    .then(function() {
        console.log('Done.');
    });