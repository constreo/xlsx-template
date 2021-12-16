import { Element } from 'elementtree';

export interface TemplatePlaceholder {
  type: string;
  string?: string;
  full: boolean;
  name: string;
  key: string;
  placeholder?: string;
}

export interface NamedTable {
  filename: string;
  root: Element;
}

interface OutputByType {
  base64: string;
  uint8array: Uint8Array;
  arraybuffer: ArrayBuffer;
  blob: Blob;
  nodebuffer: Buffer;
}

export type CellReference = {
  table?: string | null;
  colAbsolute?: boolean;
  col: string;
  colNo?: number;
  rowAbsolute?: boolean;
  row: number;
};

export type GenerateOptions = keyof OutputByType;

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

export type Placeholder = {
  full?: string;
  key?: string;
  name?: string;
  placeholder: string;
  type?: string;
  subType?: string;
};

export type Sheet = {
  id?: number;
  name?: string;
  filename: string;
  root?: Element;
};

export type SubstitutionValue = string | number | boolean | Date | Buffer;
