var stream = require('stream');
var Excel = require('exceljs');
var chai = require('chai');
var expect = chai.expect;

var headersDocument = ["Temperatura", "Humedad", "Presion atmosferica"];
var dataDocument = [[1,2,3],[4,5,6],[7,2,8],[10,2,13],[1,12,3],[15,22,31]];
var letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

getExcelOneWorkBook(headersDocument, dataDocument, "Libro1_1", "test");


function getExcelOneWorkBook(headers, data, workbookName, excelName)
{
  //Create workbook and configure
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

  //Print settings
  var worksheet =  workbook.addWorksheet(workbookName, {
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


  //Headers

  // Headers
  worksheet.getRow(rowCounter).font = {
    name: 'Arial',
    family: 4,
    size: 10,
    underline: false,
    bold: true,
  };

  worksheet.getRow(rowCounter).alignment = { vertical: 'center', horizontal: 'center' };
  worksheet.getRow(rowCounter).values =  headers;
  rowCounter++;

  //Name of webpage (Main header)
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
var maxHeaderPos = headers.length - 1;
var range = 'A1:' + letters[maxHeaderPos] + "1";
worksheet.mergeCells(range);


var rowCounter = 2;

//Header column width
for(var i=0; i<= headers.length - 1; i++)
{
  worksheet.getColumn(i + 1).width = (headers[i].length)*1.5;
}

//Data
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

//Formulas
for(var i=0; i<= headersDocument.length - 1; i++)
{
  worksheet.getCell(letters[i] + (data.length + 3)).formula === 'SUM(' +  letters[i] + "1" + ":" +  letters[i] + data.length + ')';
  console.log("Celda: " + letters[i] + (data.length + 3) + " ,Ecuacion: " + 'SUM(' +  letters[i] + "1" + ":" +  letters[i] + data.length + ')');
}

workbook.xlsx.writeFile(excelName + ".xlsx");
}