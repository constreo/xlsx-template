import sizeOf from 'buffer-image-size';
import { Element, ElementTree, parse, SubElement, tostring } from 'elementtree';
import JSZip from 'jszip';
import path from 'path';
import { CellReference, NamedTable, Options, Placeholder, Sheet, SubstitutionValue } from './types';

const DOCUMENT_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const CALC_CHAIN_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const SHARED_STRINGS_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const HYPERLINK_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

function _get_simple(obj: { [x: string]: any }, desc: string) {
  if (desc.indexOf('[') >= 0) {
    const specification = desc.split(/[[[\]]/);
    const property = specification[0];
    const index = specification[1];
    return obj[property][index];
  }

  return obj[desc];
}

/**
 * Based on http://stackoverflow.com/questions/8051975
 * Mimic https://lodash.com/docs#get
 */
function _get(obj: object, desc: string, defaultValue?: string): any {
  const arr = desc.split('.');
  try {
    while (arr.length) {
      obj = _get_simple(obj, arr.shift());
    }
  } catch (ex) {
    /* invalid chain */
    obj = undefined;
  }
  return obj === undefined ? defaultValue : obj;
}

/**
 * Create a new workbook. Either pass the raw data of a .xlsx file,
 * or call `loadTemplate()` later.
 */
export default class XlsxTemplate {
  public archive: JSZip = null;
  public sharedStrings: string[] = [];
  public sharedStringsLookup: { [key: string]: number } = {};

  protected workbook: Element = null;
  protected workbookPath: string = null;
  protected calcChainPath: string = '';

  private option: Options;
  private sharedStringsPath = '';
  private sheets: Sheet[] = [];
  private sheet: Sheet = null;
  private contentTypes: Element = null;
  private prefix: string = null;
  private workbookRels: Element = null;
  private calChainRel: Element = null;

  constructor(data: Buffer | string = null, option: Options = { imageRootPath: undefined }) {
    this.option = option;

    if (data) {
      this.loadTemplate(data);
    }
  }

  /**
   * Delete unused sheets if needed
   */
  public deleteSheet(sheetName: number): XlsxTemplate {
    const sheet = this.loadSheet(sheetName);

    const sh = this.workbook.find("sheets/sheet[@sheetId='" + sheet.id + "']");
    this.workbook.find('sheets').remove(sh);

    const rel = this.workbookRels.find("Relationship[@Id='" + sh.attrib['r:id'] + "']");
    this.workbookRels.remove(rel);

    this._rebuild();
    return this;
  }

  /**
   * Clone sheets in current workbook template
   */
  public copySheet(sheetName: number, copyName: string): XlsxTemplate {
    const sheet = this.loadSheet(sheetName); // filename, name , id, root
    const newSheetIndex = (this.workbook.findall('sheets/sheet').length + 1).toString();
    const fileName = 'worksheets' + '/' + 'sheet' + newSheetIndex + '.xml';
    const arcName = this.prefix + '/' + fileName;

    this.archive.file(arcName, tostring(sheet.root, {}));
    this.archive.file(arcName).options.binary = true;

    const newSheet = SubElement(this.workbook.find('sheets'), 'sheet');
    newSheet.attrib.name = copyName || 'Sheet' + newSheetIndex;
    newSheet.attrib.sheetId = newSheetIndex;
    newSheet.attrib['r:id'] = 'rId' + newSheetIndex;

    const newRel = SubElement(this.workbookRels, 'Relationship');
    newRel.attrib.Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    newRel.attrib.Target = fileName;

    this._rebuild();
    //    TODO: work with "definedNames"
    //    let defn = SubElement(this.workbook.find('definedNames'), 'definedName');
    //
    return this;
  }

  /**
   *  Partially rebuild after copy/delete sheets
   */
  private _rebuild() {
    // each <sheet> 'r:id' attribute in '\xl\workbook.xml'
    // must point to correct <Relationship> 'Id' in xl\_rels\workbook.xml.rels

    const order = ['worksheet', 'theme', 'styles', 'sharedStrings'];

    this.workbookRels
      .findall('*')
      .sort((rel1, rel2) => {
        // using order
        const index1 = order.indexOf(path.basename(rel1.attrib.Type));
        const index2 = order.indexOf(path.basename(rel2.attrib.Type));
        if (index1 + index2 === 0) {
          if (rel1.attrib.Id && rel2.attrib.Id)
            return parseInt(rel1.attrib.Id.substring(3), 10) - parseInt(rel2.attrib.Id.substring(3), 10);
          return (rel1 as any)._id - (rel2 as any)._id;
        }
        return index1 - index2;
      })
      .forEach((item, index) => {
        item.attrib.Id = 'rId' + (index + 1);
      });

    this.workbook.findall('sheets/sheet').forEach((item, index) => {
      item.attrib['r:id'] = 'rId' + (index + 1);
      item.attrib.sheetId = (index + 1).toString();
    });

    this.archive.file(
      this.prefix + '/' + '_rels' + '/' + path.basename(this.workbookPath) + '.rels',
      tostring(this.workbookRels, {}),
    );
    this.archive.file(this.workbookPath, tostring(this.workbook, {}));
    this.sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);
  }

  /**
   * Load a .xlsx file from a byte array.
   */
  public loadTemplate(data: Buffer | string) {
    if (Buffer.isBuffer(data)) {
      data = data.toString('binary');
    }

    this.archive = new JSZip(data, { base64: false, checkCRC32: true });

    // Load relationships
    const rels = parse(this.archive.file('_rels/.rels').asText()).getroot();
    const workbookPath = rels.find("Relationship[@Type='" + DOCUMENT_RELATIONSHIP + "']").attrib.Target;

    this.workbookPath = workbookPath;
    this.prefix = path.dirname(workbookPath);
    this.workbook = parse(this.archive.file(workbookPath).asText()).getroot();
    this.workbookRels = parse(
      this.archive.file(this.prefix + '/' + '_rels' + '/' + path.basename(workbookPath) + '.rels').asText(),
    ).getroot();
    this.sheets = this.loadSheets(this.prefix, this.workbook, this.workbookRels);
    this.calChainRel = this.workbookRels.find("Relationship[@Type='" + CALC_CHAIN_RELATIONSHIP + "']");

    if (this.calChainRel) {
      this.calcChainPath = this.prefix + '/' + this.calChainRel.attrib.Target;
    }

    this.sharedStringsPath =
      this.prefix +
      '/' +
      this.workbookRels.find("Relationship[@Type='" + SHARED_STRINGS_RELATIONSHIP + "']").attrib.Target;
    this.sharedStrings = [];
    parse(this.archive.file(this.sharedStringsPath).asText())
      .getroot()
      .findall('si')
      .forEach((si) => {
        const t = { text: '' };
        si.findall('t').forEach((tmp) => {
          t.text += tmp.text;
        });
        si.findall('r/t').forEach((tmp) => {
          t.text += tmp.text;
        });
        this.sharedStrings.push(t.text);
        this.sharedStringsLookup[t.text] = this.sharedStrings.length - 1;
      });

    this.contentTypes = parse(this.archive.file('[Content_Types].xml').asText()).getroot();
    const jpgType = this.contentTypes.find('Default[@Extension="jpg"]');
    if (jpgType === null) {
      SubElement(this.contentTypes, 'Default', { ContentType: 'image/png', Extension: 'jpg' });
    }
  }

  /**
   * Interpolate values for the sheet with the given number (1-based) or
   * name (if a string) using the given substitutions (an object).
   */
  public substitute(sheetNameOrIndex: string | number, substitutions: object) {
    const sheet = this.loadSheet(sheetNameOrIndex);
    this.sheet = sheet;

    const dimension = sheet.root.find('dimension');
    const sheetData = sheet.root.find('sheetData');
    let currentRow: number = null;
    let totalRowsInserted = 0;
    let totalColumnsInserted = 0;
    const namedTables = this.loadTables(sheet.root, sheet.filename);
    const rows = [];
    let drawing = null;

    const rels = this.loadSheetRels(sheet.filename);
    sheetData.findall('row').forEach((row) => {
      currentRow = this.getCurrentRow(row, totalRowsInserted);
      row.attrib.r = currentRow.toString();
      rows.push(row);

      const cells: Element[] = [];
      let cellsInserted = 0;
      const newTableRows: Element[] = [];

      row.findall('c').forEach((cell) => {
        let appendCell = true;
        cell.attrib.r = this.getCurrentCell(cell, currentRow, cellsInserted);

        // If c[@t="s"] (string column), look up /c/v@text as integer in
        // `this.sharedStrings`
        if (cell.attrib.t === 's') {
          // Look for a shared string that may contain placeholders
          const cellValue = cell.find('v');
          const stringIndex = parseInt(cellValue.text.toString(), 10);
          let sharedString: string = this.sharedStrings[stringIndex];

          if (sharedString === undefined) {
            return;
          }

          // Loop over placeholders
          this.extractPlaceholders(sharedString).forEach((placeholder) => {
            // Only substitute things for which we have a substitution
            let substitution: any | any[] = _get(substitutions, placeholder.name, '');
            let newCellsInserted = 0;

            if (placeholder.full && placeholder.type === 'table' && Array.isArray(substitution)) {
              if (placeholder.subType === 'image' && drawing == null) {
                if (rels) {
                  drawing = this.loadDrawing(sheet.root, sheet.filename, rels.root);
                } else {
                  // tslint:disable-next-line: no-console
                  console.log('Need to implement initRels. Or init this with Excel');
                }
              }
              newCellsInserted = this.substituteTable(
                row,
                newTableRows,
                cells,
                cell,
                namedTables,
                substitution,
                placeholder.key,
                placeholder,
                drawing,
              );

              // don't double-insert cells
              // this applies to arrays only, incorrectly applies to object arrays when there a single row, thus not rendering single row
              if (newCellsInserted !== 0 || substitution.length) {
                if (substitution.length === 1) {
                  appendCell = true;
                }
                if (substitution[0][placeholder.key] instanceof Array) {
                  appendCell = false;
                }
              }

              // Did we insert new columns (array values)?
              if (newCellsInserted !== 0) {
                cellsInserted += newCellsInserted;
                this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
              }
            } else if (placeholder.full && placeholder.type === 'normal' && Array.isArray(substitution)) {
              appendCell = false; // don't double-insert cells
              newCellsInserted = this.substituteArray(cells, cell, substitution);

              if (newCellsInserted !== 0) {
                cellsInserted += newCellsInserted;
                this.pushRight(this.workbook, sheet.root, cell.attrib.r, newCellsInserted);
              }
            } else if (placeholder.type === 'image' && placeholder.full && !Array.isArray(substitution)) {
              if (rels != null) {
                if (drawing == null) {
                  drawing = this.loadDrawing(sheet.root, sheet.filename, rels.root);
                }
                this.substituteImage(cell, sharedString, placeholder, substitution, drawing);
              } else {
                // tslint:disable-next-line: no-console
                console.log('Need to implement initRels. Or init this with Excel');
              }
            } else {
              if (placeholder.key) {
                substitution = _get(substitutions, placeholder.name + '.' + placeholder.key);
              }
              sharedString = this.substituteScalar(cell, sharedString, placeholder, (substitution instanceof Array) ? substitution[0] : substitution).toString();
            }
          });
        }

        // if we are inserting columns, we may not want to keep the original cell anymore
        if (appendCell) {
          cells.push(cell);
        }
      }); // cells loop

      // We may have inserted columns, so re-build the children of the row
      this.replaceChildren(row, cells);

      // Update row spans attribute
      if (cellsInserted !== 0) {
        this.updateRowSpan(row, cellsInserted);

        if (cellsInserted > totalColumnsInserted) {
          totalColumnsInserted = cellsInserted;
        }
      }

      // Add newly inserted rows
      if (newTableRows.length > 0) {
        // Move images for each subsitute array if option is active
        if (this.option.moveImages && rels) {
          if (drawing == null) {
            // Maybe we can load drawing at the begining of public and remove all the this.loadDrawing() along the public ?
            // If we make this, we create all the time the drawing file (like rels file at this moment)
            drawing = this.loadDrawing(sheet.root, sheet.filename, rels.root);
          }
          if (drawing != null) {
            this.moveAllImages(drawing, parseInt(row.attrib.r, 10), newTableRows.length);
          }
        }
        newTableRows.forEach((newRow) => {
          rows.push(newRow);
          ++totalRowsInserted;
        });
        this.pushDown(this.workbook, sheet.root, namedTables, currentRow, newTableRows.length);
      }
    }); // rows loop

    // We may have inserted rows, so re-build the children of the sheetData
    this.replaceChildren(sheetData, rows);

    // Update placeholders in table column headers
    this.substituteTableColumnHeaders(namedTables, substitutions);

    // Update placeholders in hyperlinks
    this.substituteHyperlinks(rels, substitutions);

    // Update <dimension /> if we added rows or columns
    if (dimension) {
      if (totalRowsInserted > 0 || totalColumnsInserted > 0) {
        const dimensionRange = this.splitRange(dimension.attrib.ref);
        const dimensionEndRef = this.splitRef(dimensionRange.end);

        dimensionEndRef.row += totalRowsInserted;
        dimensionEndRef.col = this.numToChar(this.charToNum(dimensionEndRef.col) + totalColumnsInserted);
        dimensionRange.end = this.joinRef(dimensionEndRef);

        dimension.attrib.ref = this.joinRange(dimensionRange);
      }
    }

    // Here we are forcing the values in formulas to be recalculated
    // existing as well as just substituted
    sheetData.findall('row').forEach((row) => {
      row.findall('c').forEach((cell) => {
        const formulas = cell.findall('f');
        if (formulas && formulas.length > 0) {
          cell.findall('v').forEach((v) => {
            cell.remove(v);
          });
        }
      });
    });

    // Write back the modified XML trees
    this.archive.file(sheet.filename, tostring(sheet.root, {}));
    this.archive.file(this.workbookPath, tostring(this.workbook, {}));
    if (rels) {
      this.archive.file(rels.filename, tostring(rels.root, {}));
    }
    this.archive.file('[Content_Types].xml', tostring(this.contentTypes, {}));
    // Remove calc chain - Excel will re-build, and we may have moved some formulae
    if (this.calcChainPath && this.archive.file(this.calcChainPath)) {
      this.archive.remove(this.calcChainPath);
    }

    this.writeSharedStrings();
    this.writeTables(namedTables);
    this.writeDrawing(drawing);
  }

  /**
   * Generate a new binary .xlsx file
   */
  public generate(options: JSZipGeneratorOptions): any {
    if (!options) {
      options = {
        base64: false,
      };
    }

    return this.archive.generate(options);
  }

  // Helpers

  // Write back the new shared strings list
  public writeSharedStrings(): void {
    const root = parse(this.archive.file(this.sharedStringsPath).asText()).getroot();
    const children = root.getchildren();

    root.delSlice(0, children.length);

    this.sharedStrings.forEach((sharedString) => {
      const si = Element('si');
      const t = Element('t');

      t.text = sharedString;
      si.append(t);
      root.append(si);
    });

    root.attrib.count = this.sharedStrings.length.toString();
    root.attrib.uniqueCount = this.sharedStrings.length.toString();

    this.archive.file(this.sharedStringsPath, tostring(root, {}));
  }

  // Add a new shared string
  public addSharedString(s: string): number {
    const idx = this.sharedStrings.length;
    this.sharedStrings.push(s);
    this.sharedStringsLookup[s] = idx;

    return idx;
  }

  // Get the number of a shared string, adding a new one if necessary.
  public stringIndex(s: string): number {
    let idx = this.sharedStringsLookup[s];
    if (idx === undefined) {
      idx = this.addSharedString(s);
    }
    return idx;
  }

  // Replace a shared string with a new one at the same index. Return the
  // index.
  public replaceString(oldString: string, newString: string): number {
    let idx = this.sharedStringsLookup[oldString];
    if (idx === undefined) {
      idx = this.addSharedString(newString);
    } else {
      this.sharedStrings[idx] = newString;
      delete this.sharedStringsLookup[oldString];
      this.sharedStringsLookup[newString] = idx;
    }

    return idx;
  }

  // Get a list of sheet ids, names and filenames
  protected loadSheets(prefix: string, workbook: Element, workbookRels: Element): Sheet[] {
    const sheets: Sheet[] = [];

    workbook.findall('sheets/sheet').forEach((sheet) => {
      const sheetId = sheet.attrib.sheetId;
      const relId = sheet.attrib['r:id'];
      const relationship = workbookRels.find("Relationship[@Id='" + relId + "']");
      const filename = prefix + '/' + relationship.attrib.Target;

      sheets.push({
        id: parseInt(sheetId, 10),
        name: sheet.attrib.name,
        filename,
      });
    });

    return sheets;
  }

  // Get sheet a sheet, including filename and name
  protected loadSheet(sheetNameOrIndex: string | number): Sheet {
    let info = null;

    for (const sheet of this.sheets) {
      if (
        (typeof sheetNameOrIndex === 'number' && sheet.id === sheetNameOrIndex) ||
        sheet.name === sheetNameOrIndex
      ) {
        info = sheet;
        break;
      }
    }

    if (info === null && typeof sheetNameOrIndex === 'number') {
      // Get the sheet that corresponds to the 0 based index if the id does not work
      info = this.sheets[sheetNameOrIndex - 1];
    }

    if (info === null) {
      throw new Error('Sheet ' + sheetNameOrIndex + ' not found');
    }

    return {
      filename: info.filename,
      name: info.name,
      id: info.id,
      root: parse(this.archive.file(info.filename).asText()).getroot(),
    };
  }

  // Load rels for a sheetName
  protected loadSheetRels(sheetFilename: string): Sheet {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    const relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
    const relsFile = this.archive.file(relsFilename);
    if (relsFile === null) {
      return this.initSheetRels(sheetFilename);
    }
    const rels = { filename: relsFilename, root: parse(relsFile.asText()).getroot() };
    return rels;
  }

  protected initSheetRels(sheetFilename: string): Sheet {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    const relsFilename = path.join(sheetDirectory, '_rels', sheetName + '.rels').replace(/\\/g, '/');
    const root = Element('Relationships');
    root.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
    const relsEtree = new ElementTree(root);
    const rels = { filename: relsFilename, root: relsEtree.getroot() };
    return rels;
  }

  /**
   * Load Drawing file
   */
  private loadDrawing(sheet: any, sheetFilename: string, rels: any): any {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    let drawing: any = { filename: '', root: null };
    const drawingPart = sheet.find('drawing');
    if (drawingPart === null) {
      drawing = this.initDrawing(sheet, rels);
      return drawing;
    }
    const relationshipId = drawingPart.attrib['r:id'];
    const target = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target;
    const drawingFilename = path.join(sheetDirectory, target).replace(/\\/g, '/');
    const drawingTree = parse(this.archive.file(drawingFilename).asText());
    drawing.filename = drawingFilename;
    drawing.root = drawingTree.getroot();
    drawing.relFilename = path.dirname(drawingFilename) + '/_rels/' + path.basename(drawingFilename) + '.rels';
    drawing.relRoot = parse(this.archive.file(drawing.relFilename).asText()).getroot();
    return drawing;
  }

  private addContentType(partName: string, contentType: string) {
    SubElement(this.contentTypes, 'Override', { ContentType: contentType, PartName: partName });
  }

  public initDrawing(sheet: any, rels: any): any {
    const maxId = this.findMaxId(rels, 'Relationship', 'Id', /rId(\d*)/);
    const rel = SubElement(rels, 'Relationship');
    sheet.insert(sheet._children.length, Element('drawing', { 'r:id': 'rId' + maxId }));
    rel.set('Id', 'rId' + maxId);
    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing');
    const drawing: any = {};
    const drawingFilename = 'drawing' + this.findMaxFileId(/xl\/drawings\/drawing\d*\.xml/, /drawing(\d*)\.xml/) + '.xml';
    rel.set('Target', '../drawings/' + drawingFilename);
    drawing.root = Element('xdr:wsDr');
    drawing.root.set('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
    drawing.root.set('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    drawing.filename = 'xl/drawings/' + drawingFilename;
    drawing.relFilename = 'xl/drawings/_rels/' + drawingFilename + '.rels';
    drawing.relRoot = Element('Relationships');
    drawing.relRoot.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
    this.addContentType('/' + drawing.filename, 'application/vnd.openxmlformats-officedocument.drawing+xml');
    return drawing;
  }

  /**
   * Write Drawing file
   */
  public writeDrawing(drawing: any): void {
    if (drawing !== null) {
      this.archive.file(drawing.filename, tostring(drawing.root, {}));
      this.archive.file(drawing.relFilename, tostring(drawing.relRoot, {}));
    }
  }

  /**
   * Move all images after fromRow of nbRow row
   */
  public moveAllImages(drawing: any, fromRow: number, nbRow: number): void {
    drawing.root.getchildren().forEach((drawElement: { tag: string }) => {
      if (drawElement.tag === 'xdr:twoCellAnchor') {
        this._moveTwoCellAnchor(drawElement, fromRow, nbRow);
      }
      // TODO : make the other tags image
    });
  }

  private _moveImage(drawingElement: Element, fromRow: number, nbRow: number) {
    const from = Number.parseInt(drawingElement.find('xdr:from').find('xdr:row').text.toString(), 10) + nbRow;
    drawingElement.find('xdr:from').find('xdr:row').text = from;
    const to = Number.parseInt(drawingElement.find('xdr:to').find('xdr:row').text.toString(), 10) + nbRow;
    drawingElement.find('xdr:to').find('xdr:row').text = to;
  }

  /**
   * Move TwoCellAnchor tag images after fromRow of nbRow row
   */
  private _moveTwoCellAnchor(drawingElement: any, fromRow: number, nbRow: number): void {
    if (this.option.moveSameLineImages) {
      if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text, 10) + 1 >= fromRow) {
        this._moveImage(drawingElement, fromRow, nbRow);
      }
    } else {
      if (parseInt(drawingElement.find('xdr:from').find('xdr:row').text, 10) + 1 > fromRow) {
        this._moveImage(drawingElement, fromRow, nbRow);
      }
    }
  }

  /**
   * Load tables for a given sheet
   */
  public loadTables(sheet: Element, sheetFilename: string): NamedTable[] {
    const sheetDirectory = path.dirname(sheetFilename);
    const sheetName = path.basename(sheetFilename);
    const relsFilename = sheetDirectory + '/' + '_rels' + '/' + sheetName + '.rels';
    const relsFile = this.archive.file(relsFilename);
    const tables: NamedTable[] = []; // [{filename: ..., root: ....}]

    if (relsFile === null) {
      return tables;
    }

    const rels = parse(relsFile.asText()).getroot();

    sheet.findall('tableParts/tablePart').forEach((tablePart: { attrib: { [x: string]: any } }) => {
      const relationshipId = tablePart.attrib['r:id'];
      const target = rels.find("Relationship[@Id='" + relationshipId + "']").attrib.Target;
      const tableFilename = target.replace('..', this.prefix);
      const tableTree = parse(this.archive.file(tableFilename).asText());

      tables.push({
        filename: tableFilename,
        root: tableTree.getroot(),
      });
    });

    return tables;
  }

  // Write back possibly-modified tables
  public writeTables(tables: any): void {
    tables.forEach((namedTable: { filename: string; root: Element }) => {
      this.archive.file(namedTable.filename, tostring(namedTable.root, {}));
    });
  }

  // Perform substitution in hyperlinks
  public substituteHyperlinks(rels: any, substitutions: any): void {
    parse(this.archive.file(this.sharedStringsPath).asText()).getroot();
    if (rels === null) {
      return;
    }
    const relationships = rels.root._children;
    relationships.forEach((relationship: { attrib: { Type: string; Target: string } }) => {
      if (relationship.attrib.Type === HYPERLINK_RELATIONSHIP) {
        let target = relationship.attrib.Target;

        // Double-decode due to excel double encoding url placeholders
        target = decodeURI(decodeURI(target));
        this.extractPlaceholders(target).forEach((placeholder) => {
          const substitution = substitutions[placeholder.name];

          if (substitution === undefined) {
            return;
          }
          target = target.replace(placeholder.placeholder, this.stringify(substitution));

          relationship.attrib.Target = encodeURI(target);
        });
      }
    });
  }

  // Perform substitution in table headers
  public substituteTableColumnHeaders(tables: NamedTable[], substitutions: any): void {
    tables.forEach((table) => {
      const root = table.root;
      const columns = root.find('tableColumns');
      let tableRange = this.splitRange(root.attrib.ref);
      let idx = 0;
      let inserted = 0;
      const newColumns = [];

      columns.findall('tableColumn').forEach((col: Element) => {
        ++idx;
        col.attrib.id = Number(idx).toString();
        newColumns.push(col);

        let name = col.attrib.name;

        this.extractPlaceholders(name).forEach((placeholder) => {
          const substitution = substitutions[placeholder.name];
          if (substitution === undefined) {
            return;
          }

          // Array -> new columns
          if (placeholder.full && placeholder.type === 'normal' && substitution instanceof Array) {
            substitution.forEach((element, i) => {
              let newCol = col;
              if (i > 0) {
                newCol = this.cloneElement(newCol);
                newCol.attrib.id = Number(++idx).toString();
                newColumns.push(newCol);
                ++inserted;
                tableRange.end = this.nextCol(tableRange.end);
              }
              newCol.attrib.name = this.stringify(element);
            });
            // Normal placeholder
          } else {
            name = name.replace(placeholder.placeholder, this.stringify(substitution));
            col.attrib.name = name;
          }
        });
      });

      this.replaceChildren(columns, newColumns);

      // Update range if we inserted columns
      if (inserted > 0) {
        const autoFilter = root.find('autoFilter');
        columns.attrib.count = Number(idx).toString();
        root.attrib.ref = this.joinRange(tableRange);
        if (autoFilter !== null) {
          // XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = this.joinRange(tableRange);
        }
      }

      // update ranges for totalsRowCount
      const tableRoot = table.root;
      const tableStart = this.splitRef(tableRange.start);
      const tableEnd = this.splitRef(tableRange.end);
      tableRange = this.splitRange(tableRoot.attrib.ref);

      if (tableRoot.attrib.totalsRowCount) {
        const autoFilter = tableRoot.find('autoFilter');
        if (autoFilter !== null) {
          autoFilter.attrib.ref = this.joinRange({
            start: this.joinRef(tableStart),
            end: this.joinRef(tableEnd),
          });
        }

        ++tableEnd.row;
        tableRoot.attrib.ref = this.joinRange({
          start: this.joinRef(tableStart),
          end: this.joinRef(tableEnd),
        });
      }
    });
  }

  // Return a list of tokens that may exist in the string.
  // Keys are: `placeholder` (the full placeholder, including the `${}`
  // delineators), `name` (the name part of the token), `key` (the object key
  // for `table` tokens), `full` (boolean indicating whether this placeholder
  // is the entirety of the string) and `type` (one of `table` or `cell`)
  public extractPlaceholders(sharedString: string) {
    // Yes, that's right. It's a bunch of brackets and question marks and stuff.
    const re = /\${(?:(.+?):)?(.+?)(?:\.(.+?))?(?::(.+?))??}/g;

    let match = null;
    const matches = [];
    for (match = re.exec(sharedString); match !== null; match = re.exec(sharedString)) {
      matches.push({
        placeholder: match[0],
        type: match[1] || 'normal',
        name: match[2],
        key: match[3],
        subType: match[4],
        full: match[0].length === sharedString.length,
      });
    }

    return matches;
  }

  // Split a reference into an object with keys `row` and `col` and,
  // optionally, `table`, `rowAbsolute` and `colAbsolute`.
  public splitRef(ref: string): CellReference {
    const match = ref.match(/(?:(.+)!)?(\$)?([A-Z]+)(\$)?([0-9]+)/);
    return {
      table: match?.[1] || null,
      colAbsolute: Boolean(match?.[2]),
      col: match?.[3],
      colNo: match?.[3] ? this.charToNum(match?.[3]) : null,
      rowAbsolute: Boolean(match?.[4]),
      row: parseInt(match?.[5], 10),
    };
  }

  // Join an object with keys `row` and `col` into a single reference string
  public joinRef(ref: CellReference) {
    return (
      (ref.table ? ref.table + '!' : '') +
      (ref.colAbsolute ? '$' : '') +
      ref.col.toUpperCase() +
      (ref.rowAbsolute ? '$' : '') +
      Number(ref.row).toString()
    );
  }

  // Get the next column's cell reference given a reference like "B2".
  public nextCol(ref: string) {
    ref = ref.toUpperCase();
    return ref.replace(/[A-Z]+/, (match: any) => {
      return this.numToChar(this.charToNum(match) + 1);
    });
  }

  // Get the next row's cell reference given a reference like "B2".
  public nextRow(ref: string) {
    ref = ref.toUpperCase();
    return ref.replace(/[0-9]+/, (match: string) => {
      return (parseInt(match, 10) + 1).toString();
    });
  }

  // Turn a reference like "AA" into a number like 27
  public charToNum(str: string): number {
    let num = 0;
    for (let idx = str.length - 1, iteration = 0; idx >= 0; --idx, ++iteration) {
      const thisChar = str.charCodeAt(idx) - 64; // A -> 1; B -> 2; ... Z->26
      const multiplier = Math.pow(26, iteration);
      num += multiplier * thisChar;
    }
    return num;
  }

  // Turn a number like 27 into a reference like "AA"
  public numToChar(num: number) {
    let str = '';

    for (let i = 0; num > 0; ++i) {
      const remainder = num % 26;
      let charCode = remainder + 64;
      num = (num - remainder) / 26;

      // Compensate for the fact that we don't represent zero, e.g. A = 1, Z = 26, but AA = 27
      if (remainder === 0) {
        // 26 -> Z
        charCode = 90;
        --num;
      }

      str = String.fromCharCode(charCode) + str;
    }

    return str;
  }

  // Is ref a range?
  public isRange(ref: string) {
    return ref.indexOf(':') !== -1;
  }

  // Is ref inside the table defined by startRef and endRef?
  public isWithin(ref: string, startRef: string, endRef: string) {
    const start = this.splitRef(startRef);
    const end = this.splitRef(endRef);
    const target = this.splitRef(ref);

    return start.row <= target.row && target.row <= end.row && start.colNo <= target.colNo && target.colNo <= end.colNo;
  }

  // Turn a value of any type into a string
  public stringify(value: SubstitutionValue): string {
    if (value instanceof Date) {
      // In Excel date is a number of days since 01/01/1900
      //           timestamp in ms    to days      + number of days from 1900 to 1970
      return Number(value.getTime() / (1000 * 60 * 60 * 24) + 25569).toString();
    } else if (typeof value === 'number' || typeof value === 'boolean') {
      return Number(value).toString();
    } else if (typeof value === 'string') {
      return String(value).toString();
    }

    return '';
  }

  // Insert a substitution value into a cell (c tag)
  public insertCellValue(cell: Element, substitution: SubstitutionValue): string {
    const cellValue = cell.find('v');
    const stringified = this.stringify(substitution);

    if (typeof substitution === 'string' && substitution[0] === '=') {
      // substitution, started with '=' is a formula substitution
      const formula = Element('f');
      formula.text = substitution.substr(1);
      cell.insert(1, formula);
      delete cell.attrib.t; // cellValue will be deleted later
      return formula.text.toString();
    }

    if (typeof substitution === 'number' || substitution instanceof Date) {
      delete cell.attrib.t;
      cellValue.text = stringified;
    } else if (typeof substitution === 'boolean') {
      cell.attrib.t = 'b';
      cellValue.text = stringified;
    } else {
      cell.attrib.t = 's';
      cellValue.text = Number(this.stringIndex(stringified.toString())).toString();
    }

    return stringified;
  }

  // Perform substitution of a single value
  public substituteScalar(
    cell: Element,
    str: string,
    placeholder: Placeholder,
    substitution: SubstitutionValue,
  ) {
    if (placeholder.full) {
      return this.insertCellValue(cell, substitution);
    } else {
      const newString = str.replace(placeholder.placeholder, this.stringify(substitution));
      cell.attrib.t = 's';
      return this.insertCellValue(cell, newString);
    }
  }

  // Perform a columns substitution from an array
  public substituteArray(cells: any[], cell: Element, substitution: any[]) {
    let newCellsInserted = -1; // we technically delete one before we start adding back
    let currentCell = cell.attrib.r;

    // add a cell for each element in the list
    substitution.forEach((element: any) => {
      ++newCellsInserted;

      if (newCellsInserted > 0) {
        currentCell = this.nextCol(currentCell);
      }

      const newCell = this.cloneElement(cell);
      this.insertCellValue(newCell, element);

      newCell.attrib.r = currentCell;
      cells.push(newCell);
    });

    return newCellsInserted;
  }

  // Perform a table substitution. May update `newTableRows` and `cells` and change `cell`.
  // Returns total number of new cells inserted on the original row.
  public substituteTable(
    row: Element,
    newTableRows: Element[],
    cells: Element[],
    cell: Element,
    namedTables: NamedTable[],
    substitution: object[],
    key: string,
    placeholder: Placeholder,
    drawing: any,
  ) {
    let newCellsInserted = 0; // on the original row

    // if no elements, blank the cell, but don't delete it
    if (substitution.length === 0) {
      delete cell.attrib.t;
      this.replaceChildren(cell, []);
    } else {
      const parentTables = namedTables.filter((namedTable) => {
        const range = this.splitRange(namedTable.root.attrib.ref);
        return this.isWithin(cell.attrib.r, range.start, range.end);
      });

      substitution.forEach((element: object, idx: number) => {
        let newRow: Element;
        let newCell: Element;
        let newCellsInsertedOnNewRow = 0;
        const newCells = [];
        const value = _get(element, key, '');

        if (idx === 0) {
          // insert in the row where the placeholders are

          if (value instanceof Array) {
            newCellsInserted = this.substituteArray(cells, cell, value);
          } else if (placeholder.subType === 'image' && value !== '') {
            this.substituteImage(cell, placeholder.placeholder, placeholder, value, drawing);
          } else {
            this.insertCellValue(cell, value);
          }
        } else {
          // insert new rows (or reuse rows just inserted)

          // Do we have an existing row to use? If not, create one.
          if (idx - 1 < newTableRows.length) {
            newRow = newTableRows[idx - 1];
          } else {
            newRow = this.cloneElement(row, false);
            newRow.attrib.r = this.getCurrentRow(row, newTableRows.length + 1).toString();
            newTableRows.push(newRow);
          }

          // Create a new cell
          newCell = this.cloneElement(cell, true);
          newCell.attrib.r = this.joinRef({
            row: parseInt(newRow.attrib.r, 10),
            col: this.splitRef(newCell.attrib.r).col,
          });

          if (value instanceof Array) {
            newCellsInsertedOnNewRow = this.substituteArray(newCells, newCell, value);

            // Add each of the new cells created by substituteArray()
            newCells.forEach((c) => {
              newRow.append(c);
            });

            this.updateRowSpan(newRow, newCellsInsertedOnNewRow);
          } else if (placeholder.subType === 'image' && value !== '') {
            // override fit to cell logic, since merge cells are only updated after images are replaced
            let imageDimensions;
            if (this.isMergeCell(cell)) {
              imageDimensions = this.getMergeCellDimensions(cell);
            }
            this.substituteImage(newCell, placeholder.placeholder, placeholder, value, drawing, imageDimensions);
            newRow.append(newCell);
          } else {
            this.insertCellValue(newCell, value);
            newRow.append(newCell);
          }

          // expand named table range if necessary
          parentTables.forEach((namedTable: { root: any }) => {
            const tableRoot = namedTable.root;
            const autoFilter = tableRoot.find('autoFilter');
            const range = this.splitRange(tableRoot.attrib.ref);

            if (!this.isWithin(newCell.attrib.r, range.start, range.end)) {
              range.end = this.nextRow(range.end);
              tableRoot.attrib.ref = this.joinRange(range);
              if (autoFilter !== null) {
                // XXX: This is a simplification that may stomp on some configurations
                autoFilter.attrib.ref = tableRoot.attrib.ref;
              }
            }
          });
        }
      });
    }

    return newCellsInserted;
  }

  public substituteImage(
    cell: Element,
    str: string,
    placeholder: Placeholder,
    substitution: SubstitutionValue,
    drawing: { relRoot: Element; root: Element },
    /**
     * Specific dimensions that the image should be sized to.
     * If no value is given (undefined), then the image is only fitted if the cell is a merge cell.
     */
    fitToDimensions?: { width: number, height: number },
  ) {
    this.substituteScalar(cell, str, placeholder, '');
    if (substitution === null || substitution === '') {
      return true;
    }
    // get max refid
    // update rel file.
    const maxId = this.findMaxId(drawing.relRoot, 'Relationship', 'Id', /rId(\d*)/);
    const maxFileId = this.findMaxFileId(/xl\/media\/image\d*.jpg/, /image(\d*)\.jpg/);
    const rel = SubElement(drawing.relRoot, 'Relationship');
    rel.set('Id', 'rId' + maxId);
    rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');

    rel.set('Target', '../media/image' + maxFileId + '.jpg');
    const buffer = this.imageToBuffer(substitution);
    // put image to media.
    this.archive.file('xl/media/image' + maxFileId + '.jpg', this._toArrayBuffer(buffer), {
      binary: true,
      base64: false,
    });
    const dimension = sizeOf(buffer);
    let imageWidth = this.pixelsToEMUs(dimension.width);
    let imageHeight = this.pixelsToEMUs(dimension.height);
    // let sheet = this.loadSheet(this.substitueSheetName);
    if (fitToDimensions === undefined && this.isMergeCell(cell)) {
      fitToDimensions = this.getMergeCellDimensions(cell);
    }
    if (fitToDimensions) {
      // If image is in merge cell, fit the image
      const widthRate = imageWidth / fitToDimensions.width;
      const heightRate = imageHeight / fitToDimensions.height;
      if (widthRate > heightRate) {
        imageWidth = Math.floor(imageWidth / widthRate);
        imageHeight = Math.floor(imageHeight / widthRate);
      } else {
        imageWidth = Math.floor(imageWidth / heightRate);
        imageHeight = Math.floor(imageHeight / heightRate);
      }
    } else {
      let ratio = 100;
      if (this.option && this.option.imageRatio) {
        ratio = this.option.imageRatio;
      }
      if (ratio <= 0) {
        ratio = 100;
      }
      imageWidth = Math.floor((imageWidth * ratio) / 100);
      imageHeight = Math.floor((imageHeight * ratio) / 100);
    }
    const imagePart = SubElement(drawing.root, 'xdr:oneCellAnchor');
    const fromPart = SubElement(imagePart, 'xdr:from');
    const fromCol = SubElement(fromPart, 'xdr:col');
    fromCol.text = (this.charToNum(this.splitRef(cell.attrib.r).col) - 1).toString();
    const fromColOff = SubElement(fromPart, 'xdr:colOff');
    fromColOff.text = '0';
    const fromRow = SubElement(fromPart, 'xdr:row');
    fromRow.text = (this.splitRef(cell.attrib.r).row - 1).toString();
    const fromRowOff = SubElement(fromPart, 'xdr:rowOff');
    fromRowOff.text = '0';
    const extImagePart = SubElement(imagePart, 'xdr:ext', { cx: imageWidth.toString(), cy: imageHeight.toString() });
    const picNode = SubElement(imagePart, 'xdr:pic');
    const nvPicPr = SubElement(picNode, 'xdr:nvPicPr');
    const cNvPr = SubElement(nvPicPr, 'xdr:cNvPr', { id: maxId.toString(), name: 'image_' + maxId, descr: '' });
    const cNvPicPr = SubElement(nvPicPr, 'xdr:cNvPicPr');
    const picLocks = SubElement(cNvPicPr, 'a:picLocks', { noChangeAspect: '1' });
    const blipFill = SubElement(picNode, 'xdr:blipFill');
    const blip = SubElement(blipFill, 'a:blip', {
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:embed': 'rId' + maxId,
    });
    const stretch = SubElement(blipFill, 'a:stretch');
    const fillRect = SubElement(stretch, 'a:fillRect');
    const spPr = SubElement(picNode, 'xdr:spPr');
    const xfrm = SubElement(spPr, 'a:xfrm');
    const off = SubElement(xfrm, 'a:off', { x: '0', y: '0' });
    const ext = SubElement(xfrm, 'a:ext', { cx: imageWidth.toString(), cy: imageHeight.toString() });
    const prstGeom = SubElement(spPr, 'a:prstGeom', { prst: 'rect' });
    const avLst = SubElement(prstGeom, 'a:avLst');
    const clientData = SubElement(imagePart, 'xdr:clientData');
    return true;
  }

  private isMergeCell(cell: Element) {
    for (const mergeCell of this.sheet.root.findall('mergeCells/mergeCell')) {
      // If image is in merge cell, fit the image
      if (this.cellInMergeCells(cell, mergeCell)) {
        return true;
      }
    }
    return false;
  }

  private getMergeCellDimensions(cell: Element): { width: number, height: number } {
    const mergeCell = this.sheet.root.findall('mergeCells/mergeCell').find(mc => this.cellInMergeCells(cell, mc));
    const mergeCellWidth = this.getWidthMergeCell(mergeCell, this.sheet);
    const mergeCellHeight = this.getHeightMergeCell(mergeCell, this.sheet);
    const mergeWidthEmus = this.columnWidthToEMUs(mergeCellWidth);
    const mergeHeightEmus = this.rowHeightToEMUs(mergeCellHeight);
    return {
      width: mergeWidthEmus,
      height: mergeHeightEmus,
    };
  }

  // Clone an element. If `deep` is true, recursively clone children
  public cloneElement(element: Element, deep?: boolean) {
    const newElement = Element(element.tag.toString(), element.attrib);
    newElement.text = element.text;
    newElement.tail = element.tail;

    if (deep !== false) {
      element.getchildren().forEach((child: any) => {
        newElement.append(this.cloneElement(child, deep));
      });
    }

    return newElement;
  }

  // Replace all children of `parent` with the nodes in the list `children`
  public replaceChildren(parent: Element, children: any[]) {
    parent.delSlice(0, parent.len());
    children.forEach((child: any) => {
      parent.append(child);
    });
  }

  // Calculate the current row based on a source row and a number of new rows
  // that have been inserted above
  public getCurrentRow(row: Element, rowsInserted: number): number {
    return parseInt(row.attrib.r, 10) + rowsInserted;
  }

  // Calculate the current cell based on asource cell, the current row index,
  // and a number of new cells that have been inserted so far
  public getCurrentCell(cell: Element, currentRow: number, cellsInserted: number) {
    const colRef = this.splitRef(cell.attrib.r).col;
    const colNum = this.charToNum(colRef);

    return this.joinRef({
      row: currentRow,
      col: this.numToChar(colNum + cellsInserted),
    });
  }

  // Adjust the row `spans` attribute by `cellsInserted`
  public updateRowSpan(row: Element, cellsInserted: number) {
    if (cellsInserted !== 0 && row.attrib.spans) {
      const rowSpan = row.attrib.spans.split(':').map((f: string) => parseInt(f, 10));
      rowSpan[1] += cellsInserted;
      row.attrib.spans = rowSpan.join(':');
    }
  }

  // Split a range like "A1:B1" into {start: "A1", end: "B1"}
  public splitRange(range: string) {
    const split = range.split(':');
    return {
      start: split[0],
      end: split[1],
    };
  }

  // Join into a a range like "A1:B1" an object like {start: "A1", end: "B1"}
  public joinRange(range: { start: any; end: any }) {
    return range.start + ':' + range.end;
  }

  // Look for any merged cell or named range definitions to the right of
  // `currentCell` and push right by `numCols`.
  public pushRight(workbook: Element, sheet: Element, currentCell: string, numCols: number) {
    const cellRef = this.splitRef(currentCell);
    const currentRow = cellRef.row;
    const currentCol = this.charToNum(cellRef.col);

    // Update merged cells on the same row, at a higher column
    sheet.findall('mergeCells/mergeCell').forEach((mergeCell: Element) => {
      const mergeRange = this.splitRange(mergeCell.attrib.ref);
      const mergeStart = this.splitRef(mergeRange.start);
      const mergeStartCol = this.charToNum(mergeStart.col);
      const mergeEnd = this.splitRef(mergeRange.end);
      const mergeEndCol = this.charToNum(mergeEnd.col);

      if (mergeStart.row === currentRow && currentCol < mergeStartCol) {
        mergeStart.col = this.numToChar(mergeStartCol + numCols);
        mergeEnd.col = this.numToChar(mergeEndCol + numCols);

        mergeCell.attrib.ref = this.joinRange({
          start: this.joinRef(mergeStart),
          end: this.joinRef(mergeEnd),
        });
      }
    });

    // Named cells/ranges
    workbook.findall('definedNames/definedName').forEach((name: { text: any }) => {
      const ref = name.text;

      if (this.isRange(ref)) {
        const namedRange = this.splitRange(ref);
        const namedStart = this.splitRef(namedRange.start);
        const namedStartCol = this.charToNum(namedStart.col);
        const namedEnd = this.splitRef(namedRange.end);
        const namedEndCol = this.charToNum(namedEnd.col);

        if (namedStart.row === currentRow && currentCol < namedStartCol) {
          namedStart.col = this.numToChar(namedStartCol + numCols);
          namedEnd.col = this.numToChar(namedEndCol + numCols);

          name.text = this.joinRange({
            start: this.joinRef(namedStart),
            end: this.joinRef(namedEnd),
          });
        }
      } else {
        const namedRef = this.splitRef(ref);
        const namedCol = this.charToNum(namedRef.col);

        if (namedRef.row === currentRow && currentCol < namedCol) {
          namedRef.col = this.numToChar(namedCol + numCols);

          name.text = this.joinRef(namedRef);
        }
      }
    });
  }

  // Look for any merged cell, named table or named range definitions below
  // `currentRow` and push down by `numRows` (used when rows are inserted).
  public pushDown(workbook: Element, sheet: Element, tables: NamedTable[], currentRow: number, numRows: number) {
    const mergeCells = sheet.find('mergeCells');

    // Update merged cells below this row
    sheet.findall('mergeCells/mergeCell').forEach((mergeCell: Element) => {
      const mergeRange = this.splitRange(mergeCell.attrib.ref);
      const mergeStart = this.splitRef(mergeRange.start);
      const mergeEnd = this.splitRef(mergeRange.end);

      if (mergeStart.row > currentRow) {
        mergeStart.row += numRows;
        mergeEnd.row += numRows;

        mergeCell.attrib.ref = this.joinRange({
          start: this.joinRef(mergeStart),
          end: this.joinRef(mergeEnd),
        });
      }

      // add new merge cell
      if (mergeStart.row === currentRow) {
        for (let i = 1; i <= numRows; i++) {
          const newMergeCell = this.cloneElement(mergeCell);
          mergeStart.row += 1;
          mergeEnd.row += 1;
          newMergeCell.attrib.ref = this.joinRange({
            start: this.joinRef(mergeStart),
            end: this.joinRef(mergeEnd),
          });
          mergeCells.attrib.count += 1;
          mergeCells.getchildren().push(newMergeCell);
        }
      }
    });

    // Update named tables below this row
    tables.forEach((table) => {
      const tableRoot = table.root;
      const tableRange = this.splitRange(tableRoot.attrib.ref);
      const tableStart = this.splitRef(tableRange.start);
      const tableEnd = this.splitRef(tableRange.end);

      if (tableStart.row > currentRow) {
        tableStart.row += numRows;
        tableEnd.row += numRows;

        tableRoot.attrib.ref = this.joinRange({
          start: this.joinRef(tableStart),
          end: this.joinRef(tableEnd),
        });

        const autoFilter = tableRoot.find('autoFilter');
        if (autoFilter !== null) {
          // XXX: This is a simplification that may stomp on some configurations
          autoFilter.attrib.ref = tableRoot.attrib.ref;
        }
      }
    });

    // Named cells/ranges
    workbook.findall('definedNames/definedName').forEach((name) => {
      const ref = name.text.toString();

      if (this.isRange(ref)) {
        const namedRange = this.splitRange(ref);
        const namedStart = this.splitRef(namedRange.start);
        const namedEnd = this.splitRef(namedRange.end);

        if (namedStart) {
          if (namedStart.row > currentRow) {
            namedStart.row += numRows;
            namedEnd.row += numRows;

            name.text = this.joinRange({
              start: this.joinRef(namedStart),
              end: this.joinRef(namedEnd),
            });
          }
        }
      } else {
        const namedRef = this.splitRef(ref);

        if (namedRef.row > currentRow) {
          namedRef.row += numRows;
          name.text = this.joinRef(namedRef);
        }
      }
    });
  }

  public getWidthCell(numCol: number, sheet: Sheet) {
    let defaultWidth: string = sheet.root.find('sheetFormatPr').attrib.defaultColWidth;
    if (!defaultWidth) {
      // TODO : Check why defaultColWidth is not set ?
      defaultWidth = '11.42578125';
    }
    let finalWidth = parseFloat(defaultWidth);
    sheet.root.findall('cols/col').forEach((col) => {
      if (numCol >= parseFloat(col.attrib.min) && numCol <= parseFloat(col.attrib.max)) {
        if (col.attrib.width !== undefined) {
          finalWidth = parseFloat(col.attrib.width);
        }
      }
    });
    return finalWidth;
  }

  public getWidthMergeCell(mergeCell: Element, sheet: Sheet) {
    let mergeWidth = 0;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col);
    const mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col);
    for (let i = mergeStartCol; i < mergeEndCol + 1; i++) {
      mergeWidth += this.getWidthCell(i, sheet);
    }
    return mergeWidth;
  }

  public getHeightCell(numRow: number, sheet: Sheet) {
    const defaultHight = sheet.root.find('sheetFormatPr').attrib.defaultRowHeight;
    let finalHeight = defaultHight;
    sheet.root.findall('sheetData/row').forEach((row: Element) => {
      if (numRow === parseInt(row.attrib.r, 10)) {
        if (row.attrib.ht !== undefined) {
          finalHeight = row.attrib.ht;
        }
      }
    });
    return Number.parseFloat(finalHeight);
  }

  public getHeightMergeCell(mergeCell: Element, sheet: Sheet) {
    let mergeHeight = 0;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    for (let i = mergeStartRow; i < mergeEndRow + 1; i++) {
      mergeHeight += this.getHeightCell(i, sheet);
    }
    return mergeHeight;
  }

  public getNbRowOfMergeCell(mergeCell: Element) {
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    return mergeEndRow - mergeStartRow + 1;
  }

  public pixelsToEMUs(pixels: number) {
    return Math.round((pixels * 914400) / 96);
  }

  public columnWidthToEMUs(width: number) {
    // TODO : This is not the true. Change with true calcul
    // can find help here :
    // https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
    // https://stackoverflow.com/questions/58021996/how-to-set-the-fixed-column-width-values-in-inches-apache-poi
    // https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Sheet.html#setColumnWidth-int-int-
    // https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
    // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    // http://lcorneliussen.de/raw/dashboards/ooxml/
    return this.pixelsToEMUs(width * 7.625579987895905);
  }

  public rowHeightToEMUs(height: number) {
    // TODO : need to be verify
    return Math.round((height / 72) * 914400);
  }

  protected findMaxFileId(fileNameRegex: RegExp, idRegex: RegExp) {
    const getId = (file: JSZipObject): number => {
      const filename = file.name;
      const fileIdString = idRegex.exec(filename)[1];
      return parseInt(fileIdString, 10);
    };
    const files = this.archive.file(fileNameRegex);
    const maxFileId = files.reduce((maxId: number, file) => {
      const fileId = getId(file);
      return maxId > fileId ? maxId : fileId;
    }, null);

    if (maxFileId !== null) {
      return maxFileId + 1;
    } else {
      return 1;
    }
  }

  public cellInMergeCells(cell: Element, mergeCell: Element) {
    const cellCol = this.charToNum(this.splitRef(cell.attrib.r).col);
    const cellRow = this.splitRef(cell.attrib.r).row;
    const mergeRange = this.splitRange(mergeCell.attrib.ref);
    const mergeStartCol = this.charToNum(this.splitRef(mergeRange.start).col);
    const mergeEndCol = this.charToNum(this.splitRef(mergeRange.end).col);
    const mergeStartRow = this.splitRef(mergeRange.start).row;
    const mergeEndRow = this.splitRef(mergeRange.end).row;
    if (cellCol >= mergeStartCol && cellCol <= mergeEndCol) {
      if (cellRow >= mergeStartRow && cellRow <= mergeEndRow) {
        return true;
      }
    }
    return false;
  }

  public isUrl(str: string) {
    const pattern = new RegExp(
      '^(https?:\\/\\/)?' + // protocol
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain name
        '((\\d{1,3}\\.){3}\\d{1,3}))' + // OR ip (v4) address
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // port and path
        '(\\?[;&a-z\\d%_.~+=-]*)?' + // query string
        '(\\#[-a-z\\d_]*)?$',
      'i',
    ); // fragment locator
    return !!pattern.test(str);
  }

  public toArrayBuffer(buffer: string | any[]) {
    const ab = new ArrayBuffer(buffer.length);
    const view = new Uint8Array(ab);
    for (let i = 0; i < buffer.length; ++i) {
      view[i] = buffer[i];
    }
    return ab;
  }

  public imageToBuffer(imageObj: SubstitutionValue) {
    // TODO : I think I can make this public more graceful
    if (!imageObj) {
      return null;
    }
    if (imageObj instanceof Buffer) {
      return imageObj;
    } else {
      if (typeof imageObj === 'string' || imageObj instanceof String) {
        imageObj = imageObj.toString();
        // if(this.isUrl(imageObj)){
        // TODO
        // }
        try {
          const buff = Buffer.from(imageObj, 'base64');
          return buff;
        } catch (error) {
          // tslint:disable-next-line: no-console
          console.log('this is NOT a base64 string');
          return null;
        }
      }
    }
  }

  public findMaxId(element: { findall: (arg0: any) => any[] }, tag: string, attr: string, idRegex: RegExp) {
    let maxId = 0;
    element.findall(tag).forEach((elem: Element) => {
      const match = idRegex.exec(elem.attrib[attr]);
      if (match == null) {
        throw new Error('Can not find the id!');
      }
      const cid = parseInt(match[1], 10);
      if (cid > maxId) {
        maxId = cid;
      }
    });
    return ++maxId;
  }

  private _toArrayBuffer(buffer: Buffer) {
    const ab = new ArrayBuffer(buffer.length);
    const view = new Uint8Array(ab);
    for (let i = 0; i < buffer.length; ++i) {
      view[i] = buffer[i];
    }
    return ab;
  }
}
