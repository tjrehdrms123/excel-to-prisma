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

  public async readSheet(option: IsheetOption): Promise<IrowObject[]> {
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
    return this.result;
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

  public async oneToManyCreate(option: IoneToManyCreateOption) {
    const { name, fk, rowNameIndex, startRowIndex } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = 0; i < this.result.length; i++) {
      let data: any[] = [];
      for (let j = startRowIndex; j <= sheet.rowCount; j++) {
        const rowDatas = sheet.getRow(j).values;
        const oneToOneOrManyOption: IoneToOneOrManyConnectOptions = {
          columnNames: columnNames, 
          rowDatas: rowDatas, 
          keyword: this.oneToOneOrManyConnectOptions.keyword
        };
        const rowDataObject = this.oneToOneOrManyConnect(oneToOneOrManyOption);
        if (rowDataObject[fk] !== undefined && this.result[i][fk] === rowDataObject[fk]) {
          data.push(rowDataObject);
        }
      }
      if (data.length > 0) {
        this.result[i][name] = { create: data };
      }
    }
    return option;
  }

  public async oneToManySubCreate(option: IoneToManySubCreate) {
    const { name, fk, many, rowNameIndex, startRowIndex } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = 0; i < this.result.length; i++) {
      for(let j = 0; j < this.result[i][many]["create"].length; j++) {
        let data: any[] = [];
        const manyId = this.result[i][many]["create"][j][fk];
        for (let k = startRowIndex; k <= sheet.rowCount; k++) {
          const rowDatas = sheet.getRow(k).values;
          const oneToOneOrManyOption: IoneToOneOrManyConnectOptions = {
            columnNames: columnNames, 
            rowDatas: rowDatas, 
            keyword: this.oneToOneOrManyConnectOptions.keyword
          };
          const rowDataObject = this.oneToOneOrManyConnect(oneToOneOrManyOption);
          if (rowDataObject[fk] !== undefined && manyId === rowDataObject[fk]) {
            data.push({
              ...rowDataObject,
              [fk]: manyId
            });
          }
        }
        if (data.length > 0) {
          this.result[i][many]["create"][j][name] = { create: data };
        }
      }
    }
    return option;
  }

  public getData() {
    return this.result;
  }
}