const exceljs = require('exceljs');
//Global variable 
let location = './Time Table Assistant Data.xlsx';
// To a specific cell from the sheet
function readFromCell(workbook,sheetNum,rowNum,cellNum){
    const worksheet = workbook.getWorksheet(sheetNum);
        const row = worksheet.getRow(rowNum);
        const cell = row.getCell(cellNum);
        return cell.value;
}
//To Write into a specific cell
function writeToCell(workbook,sheetNum,rowNum,cellNum,updatedValue){
    const worksheet = workbook.getWorksheet(sheetNum);
        const row = worksheet.getRow(rowNum);
        const cell = row.getCell(cellNum);
        cell.value = updatedValue;
        row.commit();//To save changes
        workbook.xlsx.writeFile(location);//To write the changes into the file
}

var workbook =new exceljs.Workbook();
workbook.xlsx.readFile(location)
.then(
    function(){
        workbook.eachSheet(function(worksheet,sheetID){ // to access each sheet in the workbook
            worksheet.eachRow(function(row,rowNumber){ // to access each row
                row.eachCell(function(cell,cellNumber){ // to access each cell
                    console.log(cell.value);
                })
            })
        })
        
        let value = readFromCell(workbook,1,2,'C');
        // console.log(value);
        // writeToCell(workbook,2,19,'A',"EVS");
        // writeToCell(workbook,2,19,'B',105);
        // writeToCell(workbook,2,19,'C',3);
        // writeToCell(workbook,2,19,'D',"LECTURE");
        // writeToCell(workbook,2,19,'E',10);
    }
)