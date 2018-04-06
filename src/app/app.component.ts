/// <reference types='@types/office-js' />
import { Component } from '@angular/core';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  output = '';
  requirements = {
    'PATIENT': {
      'required_fields': ['PATIENTID'],
      'unique_fields': ['PATIENTID'],
      'headerLineNum': 1
    },
    'SAMPLE': {
      'required_fields': ['SAMPLEID', 'PATIENTID'],
      'unique_fields': ['SAMPLEID'],
      'headerLineNum': 1
    },
    'EVENT': {
      'required_fields': ['PATIENTID', 'START', 'END'],
      'headerLineNum': 1,
      'dependencies': ['PATIENT'],
      'sheet_specific_checking': ['Type_Category_inclusion']
    },
    'GENESETS': {
      'headerLineNum': 1
    },
    'MUTATION': {
      'headerLineNum': 1,
      'dependencies': ['SAMPLE']
    },
    'MATRIX': {
      'headerLineNum': 1,
      'dependencies': ['SAMPLE']
    }
  };

  validate(): void {

    Excel.run(async (context) => {
      this.output = context.workbook.worksheets.items.toString() + '!!!';
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }
}