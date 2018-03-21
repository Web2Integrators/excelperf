import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  message;
  inprogress = false;
  perf ;
  rows = '1000';
  columns = '10';

  runCode() {
    const tableData: any[][] = this.createSampleDataRows(this.rows, this.columns);
    this.createTable(tableData);
  }

  createTable(data: any) {

    Excel.run(async (context) => {
      const t0 = performance.now();
      let isError = false;
      await OfficeHelpers.ExcelUtilities.forceCreateSheet(context.workbook, 'DD-Quarterly Net All Fees');
      const sheet = context.workbook.worksheets.getItem('DD-Quarterly Net All Fees');
      const columnLetter = this.toColumnName((data[0].length - 1) + 4);
      console.log(columnLetter);
      const example2Table = sheet.tables.add('D4:' + columnLetter + (4 + data.length), true);
      example2Table.name = 'Quarterly';
      example2Table.getDataBodyRange().values = data; // this.preppedData();

      this.message = '';
      await context.sync().catch((error) => {
        this.inprogress = false;
        isError = true;
        this.message = error;
      });
      if (!isError) {
        this.inprogress = false;
        const t1 = performance.now();
        this.perf = (t1 - t0).toFixed(0);
        this.message = `It took ${this.perf / 1000} seconds to render ${this.rows} rows`;

      }
    });
  }


  createSampleDataRows(numRows, numofColumns) {
    // tslint:disable-next-line:radix
    const a = new Array(parseInt(numRows));
    for (let i = 0; i < numRows; i++) {
      // tslint:disable-next-line:radix
      a[i] = new Array(parseInt(numofColumns));
      for (let j = 0; j < numofColumns; j++) {
        a[i][j] = '[' + i + ', ' + j + ']';
      }
    }
    return a;
  }


  toColumnName(num) {
    let ret = '';
    for (let a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      // tslint:disable-next-line:radix
      ret = String.fromCharCode(((num % b) / a) + 65) + ret;
    }
    return ret;
  }

}
