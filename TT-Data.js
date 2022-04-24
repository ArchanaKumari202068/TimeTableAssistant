const exceljs = require('exceljs');

// Requiring the module
const reader = require('xlsx') 
// Reading our TimeTableAssistantData file
const file = reader.readFile('./Time Table Assistant Data.xlsx') 
let data = []
const sheets = file.SheetNames;
// for(let i = 0; i < sheets.length; i++)
// {
//     // converting sheet data to json data
//    const sheetData = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
//    // adding json data to the array
//     data.push(sheetData)
// } 
// Printing data
// console.log(data)

// for(let i=0;i<sheets.length;i++){
//     file.Sheets[file.SheetNames[i]]
//         console.log(data)
    
// }
var workbook = new exceljs.Workbook('./Time Table Assistant Data.xlsx');
workbook.eachSheet(function(worksheet,sheetID){
    console.log(sheetID);
});   // Iterate over all sheets
// const worksheet = workbook.getWorksheet('My Sheet');  //  fetch sheet by name
const worksheet = workbook.getWorksheet(1);


// var getRowInsert = worksheet.getRow(++(lastRow.number)); // accessing row by row number


// accessing a specific cell and updating its values
    // getRowInsert.getCell('B').value = "HJFUJBJHIU";
    // getRowInsert.commit();  

// Access an individual columns by key, letter and 1-based column number
// const Teacher_id = worksheet.getColumn('B');

// iterate over all current cells in this column
// Teacher_id.eachCell(function(cell, rowNumber) {
//     // ...
//   });


// const reader = require('xlsx') 
// const TimeTableAssistantData = './TimeTableAssistantData.xlsx';
// var workbook = new exceljs.Workbook();
// workbook.xlsx.readFile(TimeTableAssistantData)
// .then(function() {
//     var worksheet = workbook.getWorksheet(2); // accessing the sheet by sheet number
//     var lastRow = worksheet.lastRow;
    
//     var getRowInsert = worksheet.getRow(++(lastRow.number)); // accessing row by row number
//     // accessing a specific cell and updating its values
//     getRowInsert.getCell('A').value = "asdf";
//     getRowInsert.commit();  
//     return workbook.xlsx.writeFile(TimeTableAssistantData)

// });



