import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs/dist/exceljs.min.js';
import * as ExcelJS from "exceljs/dist/exceljs.min.js";
import * as fs from 'file-saver';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';
declare const ExcelJS: any;
@Injectable({
  providedIn: 'root'
})
export class ExportExcelService {
constructor() {}
workbook: ExcelJS.Workbook; 
worksheet: any;

private exportExcel(json: any[], excelFileName: string, headersArray: any[]) {
 return new Promise(() => {
  //Excel Title, Header, Data
 const header = headersArray;
 const data = json;
 //Create workbook and worksheet
 this.workbook = new Workbook();
 this.worksheet = this.workbook.addWorksheet(excelFileName);
 //Add Header Row
 var headerRow = this.worksheet.addRow(header);
 // Cell Style : Fill and Border
 headerRow.eachCell((cell, number) => {
   cell.fill = {
     type: 'pattern',
     pattern: 'solid',
     fgColor: { argb: 'FF9370DB' },
     bgColor: { argb: 'FFFFFAF0' }
   }
   cell.font = { name: 'Calibri', family: 4, size: 12, bold: true };
   cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
 })
 //Add Data and Conditional Formatting
 data.forEach((element) => {
   var eachRow = [];
   headersArray.forEach((headers) => {
     eachRow.push(element[headers])
   })
   if (element.isDeleted === "Y") {
     let deletedRow = this.worksheet.addRow(eachRow);
     deletedRow.eachCell((cell, number) => {
       cell.font = { name: 'Calibri', family: 4, size: 11, bold: false, strike: true };
     })
   } else {
     this.worksheet.addRow(eachRow);
   }
 })

 this.worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber){
  
   row.eachCell(function(cell, colNumber){
     JSON.stringify(row.values);
    cell.alignment = {
      vertical: 'middle', horizontal: 'center'
    };
      for (var i = 1; i < 100; i++) {
        row.getCell(i).border = {
        top: {style:'thin'},
        left: {style:'thin'},
        bottom: {style:'thin'},
        right: {style:'thin'}
      };
    }
   
  });
 });

 this.worksheet.getColumn(1).width = 20;
 this.worksheet.getColumn(2).width = 30;
 this.worksheet.getColumn(3).width = 30;
 this.worksheet.getColumn(4).width = 30;
 this.worksheet.getColumn(4).width = 20;
 this.worksheet.getColumn(5).width = 20;
 this.worksheet.getColumn(6).width = 30;
 this.worksheet.getColumn(7).width = 30;
 this.worksheet.getColumn(8).width = 30;
 this.worksheet.getColumn(9).width = 30;
 this.worksheet.getColumn(10).width = 20;
 this.worksheet.getColumn(11).width = 20;
 this.worksheet.getColumn(12).width = 20;
 this.worksheet.getColumn(13).width = 30;
 this.worksheet.getColumn(14).width = 30;
 this.worksheet.getColumn(15).width = 20;
 this.worksheet.getColumn(16).width = 20;
 this.worksheet.getColumn(17).width = 20;
 this.worksheet.getColumn(18).width = 20;
 //this.worksheet.getCell('C').alignment = { wrapText: true };
 this.worksheet.addRow([]);

 //check column width and adjust the content accordingly
 for (let i = 0; i < this.worksheet.columns.length; i += 1) { 
   let dataMax = 0;
   const column = this.worksheet.columns[i];
   for (let j = 1; j < column.values.length; j += 1) {
     const columnLength = column.values[j].length;
     if (columnLength > dataMax) {
       //dataMax = columnLength;
       column.eachCell((cell => {
         cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
     }));
     }
   }
   column.width = dataMax < 30 ? 30 : dataMax; 
 }
 
 //Generate Excel File with given name
 this.workbook.xlsx.writeBuffer().then((data) => {
   var blob = new Blob([data], { type: EXCEL_TYPE });
   fs.saveAs(blob, excelFileName + new Date().toLocaleDateString() + EXCEL_EXTENSION);
 })
 });
}
//function for exporting file
public exportAsExcelFile(json: any[], excelFileName: string, headersArray: any[]): void {
  this.exportExcel(json, excelFileName, headersArray);
}
}