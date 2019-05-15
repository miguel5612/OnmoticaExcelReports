var stream = require('stream');
var Excel = require('exceljs');
var chai = require('chai');
var expect = chai.expect;

var headersDocument = ["Temperatura", "Humedad", "Presion atmosferica"];
var data = [[1,2,3],[4,5,6],[7,2,8],[10,2,13],[1,12,3],[15,22,31]];
var rowCounter = 2;

var workbook = new Excel.Workbook();
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Her';
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
  
  // Repeat specific rows on every printed page
  worksheet.pageSetup.printTitlesRow = '1:3';

// add a column of new values
worksheet.getRow(rowCounter).values =  headersDocument;
rowCounter++;

// display value as '1 3/5'
worksheet.getCell('A11').value = 1.6;
worksheet.getCell('A11').numFmt = '# ?/?';

// display value as '1.60%'
worksheet.getCell('B11').value = 0.016;
worksheet.getCell('B11').numFmt = '0.00%';

for(var i = 0; i<= data.length - 1; i++)
{    
    worksheet.getRow(rowCounter).values = data[i];
    rowCounter++;
}

// Iterate over all rows that have values in a worksheet
worksheet.eachRow(function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});

// Iterate over all rows (including empty rows) in a worksheet
worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
});



// for the wannabe graphic designers out there
worksheet.getCell('A1').font = {
    name: 'Comic Sans MS',
    family: 4,
    size: 16,
    underline: true,
    bold: true
};

// for the graduate graphic designers...
worksheet.getCell('A2').font = {
    name: 'Arial Black',
    color: { argb: 'FF00FF00' },
    family: 2,
    size: 14,
    italic: true
};

// for the vertical align
worksheet.getCell('A3').font = {
  vertAlign: 'superscript'
};

// note: the cell will store a reference to the font object assigned.
// If the font object is changed afterwards, the cell font will change also...
var font = { name: 'Arial', size: 12 };
worksheet.getCell('A3').font = font;
font.size = 20; // Cell A3 now has font size 20!

// Cells that share similar fonts may reference the same font object after
// the workbook is read from file or stream


// set cell alignment to top-left, middle-center, bottom-right
worksheet.getCell('A1').alignment = { vertical: 'top', horizontal: 'left' };
worksheet.getCell('B1').alignment = { vertical: 'middle', horizontal: 'center' };
worksheet.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'right' };

// set cell to wrap-text
worksheet.getCell('D1').alignment = { wrapText: true };

// set cell indent to 1
worksheet.getCell('E1').alignment = { indent: 1 };

// set cell text rotation to 30deg upwards, 45deg downwards and vertical text
worksheet.getCell('F1').alignment = { textRotation: 30 };
worksheet.getCell('G1').alignment = { textRotation: -45 };
worksheet.getCell('H1').alignment = { textRotation: 'vertical' };

//Formulas
worksheet.getCell('B3').formula === 'B1+B2';

// link to web
worksheet.getCell('A1').value = {
    text: 'Onmotica',
    hyperlink: 'http://www.onmotica.com',
    tooltip: 'WebSite'
  };
  worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };  
// merge a range of cells
var maxHeaderPos = headersDocument.length;
var letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
var range = 'A1:' + letters[maxHeaderPos] + "1";
worksheet.mergeCells(range);

  
  

  workbook.xlsx.writeFile("temp.xlsx");

  /*
// Finished adding data. Commit the worksheet
worksheet.commit();

// Finished the workbook.
workbook.commit()
  .then(function() {
    // the stream has been written
  });*/