var stream = require('stream');
var Excel = require('exceljs');
var chai = require('chai');
var expect = chai.expect;

var headersDocument = ["Temperatura", "Humedad", "Presion atmosferica"];
var data = [[1,2,3],[4,5,6],[7,2,8],[10,2,13],[1,12,3],[15,22,31]];
var rowCounter = 2;

var workbook = new Excel.Workbook();
workbook.creator = 'Onmotica';
workbook.lastModifiedBy = 'Onmotica';
workbook.created = new Date();
workbook.modified = new Date();
workbook.lastPrinted = new Date();
// Set workbook dates to 1904 date system
workbook.properties.date1904 = true;
workbook.views = [
  {
    x: 0, y: 0, width: 10000, height: 20000,
    firstSheet: 0, activeTab: 1, visibility: 'visible'
  }
]

//var sheet = workbook.addWorksheet('My Sheet');
var worksheet =  workbook.addWorksheet('sheet', {
  pageSetup:{paperSize: 9, orientation:'landscape'}
});
  
 
// adjust pageSetup settings afterwards
worksheet.pageSetup.margins = {
  left: 0.7, right: 0.7,
  top: 0.75, bottom: 0.75,
  header: 0.3, footer: 0.3
};

// Set Print Area for a sheet
worksheet.pageSetup.printArea = 'A1:G20';

// Headers
worksheet.getRow(rowCounter).font = {
  name: 'Arial',
  family: 4,
  size: 10,
  underline: false,
  bold: true,
};

worksheet.getRow(rowCounter).alignment = { vertical: 'center', horizontal: 'center' };
worksheet.getRow(rowCounter).values =  headersDocument;
rowCounter++;
//Header column width
for(var i=0; i<= headersDocument.length - 1; i++)
{
  worksheet.getColumn(i + 1).width = (headersDocument[i].length)*1.5;
}

for(var i = 0; i<= data.length - 1; i++)
{    
    // for the wannabe graphic designers out there
    worksheet.getRow(rowCounter).font = {
      name: 'Arial',
      family: 4,
      size: 8,
      underline: false,
      bold: false
    };
    worksheet.getRow(rowCounter).values = data[i];
    worksheet.getRow(rowCounter).alignment = { vertical: 'center', horizontal: 'center' };
    rowCounter++;
}

// Iterate over all rows that have values in a worksheet
worksheet.eachRow(function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// link to web
worksheet.getCell('A1').value = {
    text: 'Onmotica',
    hyperlink: 'http://www.onmotica.com',
    tooltip: 'WebSite'
  };
  worksheet.getCell('A1').font = {
    name: 'Arial',
    family: 4,
    size: 12,
    underline: false,
    bold: false
  };
  worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };  
// merge a range of cells
var maxHeaderPos = headersDocument.length - 1;
var letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
var range = 'A1:' + letters[maxHeaderPos] + "1";
worksheet.mergeCells(range);

//Formulas
for(var i=0; i<= headersDocument.length - 1; i++)
{
  worksheet.getCell(letters[i] + (data.length + 3)).formula === 'SUM(' +  letters[i] + "1" + ":" +  letters[i] + data.length + ')';
  console.log("Celda: " + letters[i] + (data.length + 3) + " ,Ecuacion: " + 'SUM(' +  letters[i] + "1" + ":" +  letters[i] + data.length + ')');
}

workbook.xlsx.writeFile("temp.xlsx");