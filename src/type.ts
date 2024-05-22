export type ToneToOneOrManyOperationType = "connect" | "set" | "disconnect"

export type ToneToOneOrManyOptionType = "one" | "many"

export type TrowObject = {
  [key: string]: any;
}

export type TsheetOption = {
  name: string;
  rowNameIndex: number;
  startRowIndex: number;
  oneToOneOrManyOptions?: ToneToOneOrManyOption[];
}

export type TsplitKeyword = {
  split: string;
}

export type TconstructorOptions = {
  filePath: string;
  pkDelimiterString: string;
  oneToOneOrManyOptions: TsplitKeyword;
}

export type ToneToOneOrManyParse = {
  value: string | number;
  optionType: ToneToOneOrManyOptionType;
}

export type ToneToOneOrManyOptions = {
  columnNames: any,
  rowDatas: any,
  oneToOneOrManyOptions?: ToneToOneOrManyOption[];
}

export type ToneToManySubCreate = {
  name: string;
  fk: string;
  many: string;
  rowNameIndex: number;
  startRowIndex: number;
  oneToOneOrManyOptions?: ToneToOneOrManyOption[];
}

export type TparseCreate = {
  obj: TrowObject,
  newData: TrowObject,
  fk: string,
  name: string,
  many: string
}

export type ToneToOneOrManyOption = {
  key: string;
  option: ToneToOneOrManyOptionType;
  operation: ToneToOneOrManyOperationType;
}