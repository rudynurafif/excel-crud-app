import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { ExcelData } from 'src/app/models/excel-data.model';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  private excelData: ExcelData[] = [
    { employeeId: 1, name: 'John', age: 30 },
    { employeeId: 2, name: 'Jane', age: 25 },
  ];

  constructor() {}

  getExcelData(): ExcelData[] {
    return this.excelData;
  }

  addRow(newRow: ExcelData) {
    newRow.employeeId = this.excelData.length + 1;
    this.excelData.push(newRow);
  }

  editRow(updatedRow: ExcelData) {
    const index = this.excelData.findIndex(row => row.employeeId === updatedRow.employeeId);
    if (index !== -1) {
      this.excelData[index] = updatedRow;
    }
  }

  deleteRow(rowemployeeId: number) {
    const index = this.excelData.findIndex(row => row.employeeId === rowemployeeId);
    if (index !== -1) {
      this.excelData.splice(index, 1);
    }
  }

  exportToExcel() {
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.excelData);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'excel-data.xlsx');
  }
}
