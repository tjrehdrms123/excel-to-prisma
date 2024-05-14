/**
 * Naming principle
 * I: interface
 * C: constructor
 */

export type IRowObject = {
  [key: string]: any;
}

export type ISheetOption = {
  name: string;
  rowNameIndex: number;
  startRowIndex: number;
}

export type ISheetManyToManyOption = {
  name: string;
  fk: string;
  rowNameIndex: number;
  startRowIndex: number;
}

export type IConeToOneOrManyOptions = {
  keyword: string;
  split: string;
}

export type IconstructorOptions = {
  filePath: string;
  oneToOneOrManyOptions: IConeToOneOrManyOptions;
}

export type IoneToOneOrManyOptions = {
  columnNames: any,
  rowDatas: any,
  keyword: string
}

export type IparseConnect = {
  value: any;
}

export type ImanyToManySub = {
  name: string;
  fk: string;
  many: string;
  rowNameIndex: number;
  startRowIndex: number;
}