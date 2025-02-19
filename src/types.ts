import { Element } from 'elementtree';

export interface TemplatePlaceholder {
  type: string;
  string?: string;
  full: boolean;
  name: string;
  key: string;
  placeholder?: string;
  subType?: string;
}

export interface NamedTable {
  filename: string;
  root: Element;
}

export type GenerateOptions = {
  base64?: boolean;
  compression?: string;
  /** base64 (default), string, uint8array, blob */
  type?: string;
  comment?: string;
}

export type CellReference = {
  table?: string | null;
  colAbsolute?: boolean;
  col: string;
  colNo?: number;
  rowAbsolute?: boolean;
  row: number;
};

export interface RangeSplit {
  start: string;
  end: string;
}

export interface ReferenceAddress {
  table?: string;
  colAbsolute?: boolean;
  col: string;
  rowAbsolute?: boolean;
  row: number;
}

export type Options = {
  imageRootPath?: string;
  moveImages?: boolean;
  imageRatio?: number;
  moveSameLineImages?: boolean;
};

export type Sheet = {
  id?: number;
  name?: string;
  filename: string;
  root?: Element;
};

export type SubstitutionValue = string | number | boolean | Date | Buffer;
