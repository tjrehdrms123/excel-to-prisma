import exceljs from 'exceljs';
import { IRowObject, ISheetOption, ISheetManyToManyOption, IoneToOneOrManyOptions, IconstructorOptions, IConeToOneOrManyOptions, IparseConnect, ImanyToManySub } from './interface';

export class ExcelToPrisma {
  private workbook: any;
  private result: IRowObject[] = [];
  private filePath: string;
  private oneToOneOrManyOptions: IConeToOneOrManyOptions;

  constructor(options: IconstructorOptions) {
    this.workbook = new exceljs.Workbook();
    this.filePath = options.filePath;
    this.oneToOneOrManyOptions = options.oneToOneOrManyOptions;
  }

  public async initialize(): Promise<void> {
    await this.workbook.xlsx.readFile(this.filePath);
  }

  public async readSheet(sheetOption: ISheetOption): Promise<IRowObject[]> {
    const { name, rowNameIndex, startRowIndex } = sheetOption;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = startRowIndex; i <= sheet.rowCount; i++) {
      const rowDatas = sheet.getRow(i).values;
      const oneToOneOrManyOption: IoneToOneOrManyOptions = {
        columnNames: columnNames, 
        rowDatas: rowDatas, 
        keyword: this.oneToOneOrManyOptions.keyword
      };
      const rowDataObject = this.oneToOneOrMany(oneToOneOrManyOption);
      if (rowDataObject[`${name}Id`] !== undefined) {
        this.result.push(rowDataObject);
      }
    }
    return this.result;
  }

  private oneToOneOrMany(sheetOption: IoneToOneOrManyOptions) {
    const { columnNames, rowDatas, keyword } = sheetOption;
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
        : value.split(this.oneToOneOrManyOptions.split).map((id: string) => ({ id: parseInt(id) }));
    }
  }

  public async manyToMany(sheetManyToManyOption: ISheetManyToManyOption) {
    const { name, fk, rowNameIndex, startRowIndex } = sheetManyToManyOption;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = 0; i < this.result.length; i++) {
      let data: any[] = [];
      for (let j = startRowIndex; j <= sheet.rowCount; j++) {
        const rowDatas = sheet.getRow(j).values;
        const oneToOneOrManyOption: IoneToOneOrManyOptions = {
          columnNames: columnNames, 
          rowDatas: rowDatas, 
          keyword: this.oneToOneOrManyOptions.keyword
        };
        const rowDataObject = this.oneToOneOrMany(oneToOneOrManyOption);
        if (rowDataObject[fk] !== undefined && this.result[i][fk] === rowDataObject[fk]) {
          data.push(rowDataObject);
        }
      }
      if (data.length > 0) {
        this.result[i][name] = { create: data };
      }
    }
    return sheetManyToManyOption;
  }

  public async manyToManySub(sheetOption: ImanyToManySub) {
    const { name, fk, many, rowNameIndex, startRowIndex } = sheetOption;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = 0; i < this.result.length; i++) {
      for(let j = 0; j < this.result[i][many]["create"].length; j++) {
        let data: any[] = [];
        const manyId = this.result[i][many]["create"][j][fk];
        for (let k = startRowIndex; k <= sheet.rowCount; k++) {
          const rowDatas = sheet.getRow(k).values;
          const oneToOneOrManyOption: IoneToOneOrManyOptions = {
            columnNames: columnNames, 
            rowDatas: rowDatas, 
            keyword: this.oneToOneOrManyOptions.keyword
          };
          const rowDataObject = this.oneToOneOrMany(oneToOneOrManyOption);
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
    return sheetOption;
  }

  public getData() {
    return this.result;
  }
}