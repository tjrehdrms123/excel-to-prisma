/**
 * Naming principle
 * I: interface
 * C: constructor
 */

export type IrowObject = {
  [key: string]: any;
}

export type IsheetOption = {
  name: string;
  rowNameIndex: number;
  startRowIndex: number;
}

export type IoneToManyCreateOption = {
  name: string;
  fk: string;
  rowNameIndex: number;
  startRowIndex: number;
}

export type IConeToOneOrManyConnectOptions = {
  keyword: string;
  split: string;
}

export type IconstructorOptions = {
  filePath: string;
  oneToOneOrManyConnectOptions: IConeToOneOrManyConnectOptions;
}

export type IoneToOneOrManyConnectOptions = {
  columnNames: any,
  rowDatas: any,
  keyword: string
}

export type IparseConnect = {
  value: any;
}

export type IoneToManySubCreate = {
  name: string;
  fk: string;
  many: string;
  rowNameIndex: number;
  startRowIndex: number;
}