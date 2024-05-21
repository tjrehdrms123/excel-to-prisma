/**
 * Naming principle
 * I: interface
 * C: constructor
 * E: enum
 */

export enum EoneToOneOrManyOperation {
  CONNECT = 'connect',
  SET = 'set',
  DISCONNECT = 'disconnect'
}

export type IrowObject = {
  [key: string]: any;
}

export type IsheetOption = {
  name: string;
  rowNameIndex: number;
  startRowIndex: number;
  oneKeyword?: Array<string>;
  manyKeyword?: Array<string>;
  oneToOneOrManyOperation: EoneToOneOrManyOperation;
}

export type IConeToOneOrManyConnectOptions = {
  split: string;
}

export type IconstructorOptions = {
  filePath: string;
  oneToOneOrManyConnectOptions: IConeToOneOrManyConnectOptions;
}

export type IoneToOneOrManyConnectOptions = {
  columnNames: any,
  rowDatas: any,
  oneKeyword?: Array<string>;
  manyKeyword?: Array<string>;
  oneToOneOrManyOperation: EoneToOneOrManyOperation;
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
  oneKeyword?: Array<string>;
  manyKeyword?: Array<string>;
  oneToOneOrManyOperation: EoneToOneOrManyOperation;
}

export type IparseCreate = {
  obj: IrowObject,
  newData: IrowObject,
  fk: string,
  name: string,
  many: string
}