import exceljs from 'exceljs';
import { TrowObject, TsheetOption, TsplitKeyword, TconstructorOptions, ToneToManySubCreate, ToneToOneOrManyOptions, TparseCreate, ToneToOneOrManyParse } from './type';

export class ExcelToPrisma {
  private workbook: any;
  private result: TrowObject[] = [];
  private filePath: string;
  private oneToOneOrManyOptions: TsplitKeyword;
  private pkDelimiterString: string;

  constructor(options: TconstructorOptions) {
    this.workbook = new exceljs.Workbook();
    this.filePath = options.filePath;
    this.oneToOneOrManyOptions = options.oneToOneOrManyOptions;
    this.pkDelimiterString = options.pkDelimiterString;
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
  public async readSheet(option: TsheetOption) {
    const { name, rowNameIndex, startRowIndex, oneToOneOrManyOptions } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
    for (let i = startRowIndex; i <= sheet.rowCount; i++) {
      const rowDatas = sheet.getRow(i).values;
      const oneToOneOrManyOption: ToneToOneOrManyOptions = {
        columnNames: columnNames, 
        rowDatas: rowDatas, 
        oneToOneOrManyOptions: oneToOneOrManyOptions
      };
      const rowDataObject = this.oneToOneOrMany(oneToOneOrManyOption);
      // Exclude missing data from sheet
      if (rowDataObject[`${name}${this.pkDelimiterString}`] !== undefined) {
        this.result.push(rowDataObject);
      }
    }
    return option;
  }

  /** 
   * @description Convert data to prisma format with oneToOneOrMany function
   * @param option List of options to setting data parsing
  */
  private oneToOneOrMany(option: ToneToOneOrManyOptions) {
    const { columnNames, rowDatas, oneToOneOrManyOptions } = option;
    const obj: { [key: string]: any } = {};
    columnNames.forEach((columnName: string, index: number) => {
      obj[columnName] = rowDatas[index];
      // If the oneToOneOrManyOptions option exists, parse it in oneToOneOrManyOptions.option format.
      if(oneToOneOrManyOptions != undefined && oneToOneOrManyOptions.length > 0) {
        oneToOneOrManyOptions.forEach((oneToOneOrManyOption) => {
          const { key, option, operation } = oneToOneOrManyOption;
          if (columnName === key && rowDatas[index] !== undefined) {
            obj[columnName] = {
              [operation]: this.oneToOneOrManyParse({ value: rowDatas[index], optionType: option })
            };
          }
        });
      }
    });
    return obj;
  }

  /** 
   * @description Convert data to one or many prisma format
   * @param value target value to convert
  */
  private oneToOneOrManyParse(option: ToneToOneOrManyParse) {
    const { value, optionType } = option;
    if (value !== undefined) {
      let data;
      if(optionType === 'one') {
        data = { id: parseInt(value.toString()) };
      }
      if (optionType === 'many') {
        if(typeof value === 'number') {
          data = [{ id: parseInt(value.toString()) }]
        } else {
          data = value.split(this.oneToOneOrManyOptions.split).map((id: string) => ({ id: parseInt(id) }));
        }
      }
      return data;
    }
  }

  /** 
   * @description Convert data to prisma-create format with parseCreate function
   * @param option List of options to setting data parsing
  */
  public async oneToManyCreate(option: ToneToManySubCreate) {
    const { name, fk, rowNameIndex, startRowIndex, many, oneToOneOrManyOptions } = option;
    const sheet = this.workbook.getWorksheet(name);
    const columnNames = sheet.getRow(rowNameIndex).values;
  
    const obj = [];
  
    for (let j = startRowIndex; j <= sheet.rowCount; j++) {
      const rowDatas = sheet.getRow(j).values;
      const oneToOneOrManyOption: ToneToOneOrManyOptions = {
        columnNames: columnNames, 
        rowDatas: rowDatas, 
        oneToOneOrManyOptions: oneToOneOrManyOptions
      };
      const rowDataObject = this.oneToOneOrMany(oneToOneOrManyOption);
      obj.push(rowDataObject);
    }

    for (const newData of obj) {
      this.parseCreate({ obj: this.result, newData: newData, fk: fk, name: name, many: many});    
    }
    return option;
  }
  
  /** 
   * @description Convert the data to prisma-create format by recursively calling the values.
   * @param option List of options to setting data parsing
  */
  private parseCreate(option: TparseCreate): boolean {
    const { obj, newData, fk, name, many } = option;
    if (Array.isArray(obj)) {
      // Add data by recursively calling a method for each item in the array
      for (const item of obj) {
        if (this.parseCreate({obj: item, newData: newData, fk: fk, name: name, many: many })) return true;
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
          if (this.parseCreate({ obj: obj[key], newData: newData, fk: fk, name: name, many: many })) return true;
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
  private removeEmptyArrays(data: Array<TrowObject>): Array<TrowObject> {
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
  public getData(): Array<TrowObject> {
    return this.removeEmptyArrays(this.result);
  }
}