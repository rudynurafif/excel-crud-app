import { Component } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { ExcelData } from 'src/app/models/excel-data.model';
import { ExcelService } from 'src/app/services/excel/excel.service';

@Component({
  selector: 'app-excel',
  templateUrl: './excel.component.html',
  styleUrls: ['./excel.component.scss']
})
export class ExcelComponent {

  excelData: ExcelData[] = [];
  excelForm!: FormGroup;

  constructor(private excelService: ExcelService, private formBuilder: FormBuilder) { }

  ngOnInit(): void {
    this.loadExcelData();
    this.excelForm = this.formBuilder.group({
      name: ['', Validators.required],
      age: ['', Validators.required],
    });
  }

  loadExcelData() {
    this.excelData = this.excelService.getExcelData();
  }

  addRow() {
    if (this.excelForm.valid) {
      const newEntry: ExcelData = this.excelForm.value;
      this.excelService.addRow(newEntry);
      this.loadExcelData();
      this.excelForm.reset();
    }
  }

  editRow(rowIndex: number) {
  }

  deleteRow(rowIndex: number) {
    this.excelService.deleteRow(rowIndex);
    this.loadExcelData();
  }

  exportToExcel() {
    this.excelService.exportToExcel();
  }

}
