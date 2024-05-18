import exceljs from 'exceljs';
import { IrowObject, IsheetOption, IConeToOneOrManyConnectOptions, IconstructorOptions, IoneToManySubCreate, IparseConnect, IoneToOneOrManyConnectOptions, IparseCreate } from './interface';

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

  /**
   * @description Initialize Excel file
  */
  public async initialize(): Promise<void> {
    await this.workbook.xlsx.readFile(this.filePath);
  }

  /** 
   * @description Read parent table data into Excel file
   * @param option List of options to setting data parsing
  */
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
      // Exclude missing data from sheet 
      if (rowDataObject[`${name}Id`] !== undefined) {
        this.result.push(rowDataObject);
      }
    }
    return option;
  }

  /** 
   * @description Convert data to prisma-connect format with parseConnect function
   * @param option List of options to setting data parsing
  */
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

  /** 
   * @description Convert data to one or many prisma-connect format
   * @param value target value to convert
  */
  private parseConnect(value: any) {
    if (value !== undefined && value !== false) {
      return typeof value === "number"
        ? [{ id: parseInt(value.toString()) }]
        : value.split(this.oneToOneOrManyConnectOptions.split).map((id: string) => ({ id: parseInt(id) }));
    }
  }

  /** 
   * @description Convert data to prisma-create format with parseCreate function
   * @param option List of options to setting data parsing
  */
  public async oneToManyCreate(option: IoneToManySubCreate) {
    const { name, fk, rowNameIndex, startRowIndex, many } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
  
    const obj = [];
  
    for (let j = startRowIndex; j <= sheet.rowCount; j++) {
      const rowDatas = sheet.getRow(j).values;
      const oneToOneOrManyOption: IoneToOneOrManyConnectOptions = {
        columnNames: columnNames,
        rowDatas: rowDatas,
        keyword: this.oneToOneOrManyConnectOptions.keyword
      };
      const rowDataObject = this.oneToOneOrManyConnect(oneToOneOrManyOption);
      obj.push(rowDataObject);
    }

    for (const newData of obj) {
      this.parseCreate({
        obj: this.result,
        newData: newData,
        fk: fk,
        name: name,
        many: many});
    }
  
    return option;
  }
  
  /** 
   * @description Convert the data to prisma-create format by recursively calling the values.
   * @param option List of options to setting data parsing
  */
  private parseCreate(option: IparseCreate): boolean {
    const { obj, newData, fk, name, many } = option;
    if (Array.isArray(obj)) {
      // Add data by recursively calling a method for each item in the array
      for (const item of obj) {
        if (this.parseCreate({
          obj: item,
          newData: newData,
          fk: fk,
          name: name,
          many: many
        })) return true;
      }
    } else if (typeof obj === 'object' && obj !== null) {
      // Check whether the foreign key (fk) of obj matches the foreign key of the new object (newData)
      if (obj[fk] === newData[fk]) {
        // If the object (obj) does not have a name property, add the name property and create an array
        if (!obj[name]) {
          obj[name] = { create: [] };
        }
        if (newData[fk] != undefined) {
          obj[name].create.push(newData);
        }
        return true;
      }
      // Add data by recursively calling methods for each property of an object.
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          if (this.parseCreate({
            obj: obj[key],
            newData: newData,
            fk: fk,
            name: name,
            many: many
          })) return true;
        }
      }
    }
    return false;
  }

  /**
   * @description Remove empty arrays
   * @todo Empty array needs to be removed from parseCreate method
   * @param data target data to remove empty arrays
   */
  private removeEmptyArrays(data: Array<IrowObject>): Array<IrowObject> {
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

  /**
   * @description Get data
   */
  public getData(): Array<IrowObject> {
    return this.removeEmptyArrays(this.result);
  }
}