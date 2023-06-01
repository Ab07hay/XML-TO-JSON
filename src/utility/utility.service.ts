import { Injectable } from '@nestjs/common';
import { parts } from '../JsonFiles/parts';
import { partAbbr } from '../JsonFiles/part-abbr';
import * as ExcelJS from 'exceljs';
@Injectable()
export class UtilityService {
  async getPartsPositionRule() {
    const workbook = new ExcelJS.Workbook();
    const filePath = 'src/file.xlsx';
    await workbook.xlsx.readFile(filePath);
    const jsonData = [];
    const rules = {};
    const headers = [];
    workbook.worksheets.forEach((worksheet) => {
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const rowData = {};
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (rowNumber === 1) {
            headers[colNumber] = cell.value;
          } else {
            rowData[headers[colNumber]] = cell.value;
          }
        });
        if (rowNumber !== 1) {
          jsonData.push(rowData);
        }
      });
    });

    jsonData.forEach((p) => {
      const partsKeys = ['First', 'Second', 'Third', 'Fourth', 'Fifth'];
      const partName = partsKeys
        .map((key) => p[key])
        .filter((part) => parts.includes(part))
        .sort()
        .map((part) => partAbbr[part])
        .join('-');

      rules[partName] = p.Position;
    });

    return jsonData;
  }
}
