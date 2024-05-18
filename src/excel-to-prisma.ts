import exceljs from 'exceljs';
import { IrowObject, IsheetOption, IConeToOneOrManyConnectOptions, IoneToManyCreateOption, IconstructorOptions, IoneToManySubCreate, IparseConnect, IoneToOneOrManyConnectOptions } from './interface';

export class ExcelToPrisma {
  private workbook: any;
  private result: IrowObject[] = [];
  private filePath: string;
  private oneToOneOrManyConnectOptions: IConeToOneOrManyConnectOptions;

  constructor(options: IconstructorOptions) {
    this.workbook = new exceljs.Workbook();
    this.filePath = options.filePath;
    this.oneToOneOrManyConnectOptions = options.oneToOneOrManyConnectOptions;
  }

  public async initialize(): Promise<void> {
    await this.workbook.xlsx.readFile(this.filePath);
  }

  public async readSheet(option: IsheetOption) {
    const { name, rowNameIndex, startRowIndex } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = startRowIndex; i <= sheet.rowCount; i++) {
      const rowDatas = sheet.getRow(i).values;
      const oneToOneOrManyOption: IoneToOneOrManyConnectOptions = {
        columnNames: columnNames, 
        rowDatas: rowDatas, 
        keyword: this.oneToOneOrManyConnectOptions.keyword
      };
      const rowDataObject = this.oneToOneOrManyConnect(oneToOneOrManyOption);
      if (rowDataObject[`${name}Id`] !== undefined) {
        this.result.push(rowDataObject);
      }
    }
    return option;
  }

  private oneToOneOrManyConnect(option: IoneToOneOrManyConnectOptions) {
    const { columnNames, rowDatas, keyword } = option;
    const obj: { [key: string]: any } = {};
    columnNames.forEach((columnName: string, index: number) => {
      obj[columnName] = rowDatas[index];
      if (columnName.length > keyword.length && columnName.substring(0, keyword.length) === keyword && rowDatas[index] !== undefined && rowDatas[index] !== false) {
        obj[columnName] = {
          connect: this.parseConnect(rowDatas[index]),
        };
      }
    });

    return obj;
  }

  private parseConnect(value: any) {
    if (value !== undefined && value !== false) {
      return typeof value === "number"
        ? [{ id: parseInt(value.toString()) }]
        : value.split(this.oneToOneOrManyConnectOptions.split).map((id: string) => ({ id: parseInt(id) }));
    }
  }

  public async oneToManyCreate(option: IoneToManySubCreate) {
    const { name, fk, rowNameIndex, startRowIndex, many } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
  
    const dataToAdd = [];
  
    for (let j = startRowIndex; j <= sheet.rowCount; j++) {
      const rowDatas = sheet.getRow(j).values;
      const oneToOneOrManyOption: IoneToOneOrManyConnectOptions = {
        columnNames: columnNames,
        rowDatas: rowDatas,
        keyword: this.oneToOneOrManyConnectOptions.keyword
      };
      const rowDataObject = this.oneToOneOrManyConnect(oneToOneOrManyOption);
      dataToAdd.push(rowDataObject);
    }

    for (const item of dataToAdd) {
      this.addCommentHistory(this.result, item, fk, name, many);
    }
  
    return option;
  }
  
  private addCommentHistory(obj: any, newHistory: any, fk: string, name: string, many: string): boolean {
    if (Array.isArray(obj)) {
      // 배열의 각 항목에 대해 재귀적으로 함수를 호출하여 데이터를 추가
      for (const item of obj) {
        if (this.addCommentHistory(item, newHistory, fk, name, many)) return true;
      }
    } else if (typeof obj === 'object' && obj !== null) {
      // obj의 외래 키(fk)가 새로운 객체(newHistory)의 외래 키와 일치하는지 확인
      if (obj[fk] === newHistory[fk]) {
        // 만약 객체(obj)에 name 속성이 없는 경우, name 속성을 추가하고 create 배열을 생성
        if (!obj[name]) {
          obj[name] = { create: [] };
        }
        if (newHistory[fk] != undefined) {
          obj[name].create.push(newHistory);
        }
        return true;
      }
      // 객체의 각 속성에 대해 재귀적으로 함수를 호출하여 데이터를 추가
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          if (this.addCommentHistory(obj[key], newHistory, fk, name, many)) return true;
        }
      }
    }
    return false;
  }

  private removeEmptyArrays(data: any[]): any[] {
    return data.map(user => {
        for (const key in user) {
            if (Array.isArray(user[key]) && user[key].length === 0) {
                delete user[key];
            } else if (typeof user[key] === 'object' && user[key] !== null) {
                user[key] = this.removeEmptyArrays([user[key]])[0];
            }
        }
        return user;
    }).filter(user => Object.keys(user).length > 0);
}


  public getData() {
    return this.removeEmptyArrays(this.result);
  }
}