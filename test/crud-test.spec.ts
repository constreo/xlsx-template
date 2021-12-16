import etree from 'elementtree';
import fs from 'fs';
import path from 'path';
import XlsxTemplate from '../src';

function getSharedString(sharedStrings, sheet1, index) {
  return sharedStrings.findall("./si")[
    parseInt(sheet1.find("./sheetData/row/c[@r='" + index + "']/v").text.toString(), 10)
  ].find("t").text.toString();
}

describe("CRUD operations", () => {

  describe('XlsxTemplate', () => {

    it("can load data", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);
        expect(t.sharedStrings).toEqual([
          "Name", "Role", "Plan table", "${table:planData.name}",
          "${table:planData.role}", "${table:planData.days}",
          "${dates}", "${revision}",
          "Extracted on ${extractDate}"
        ]);

        done();
      });

    });

    it("can write changed shared strings", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.replaceString("Plan table", "The plan");

        t.writeSharedStrings();

        const text = t.archive.file("xl/sharedStrings.xml").asText();
        expect(text).not.toMatch("<si><t>Plan table</t></si>");
        expect(text).toMatch("<si><t>The plan</t></si>");

        done();
      });

    });

    it("can substitute values and generate a file", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          extractDate: new Date("2013-01-02"),
          revision: 10,
          dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
          planData: [
            {
              name: "John Smith",
              role: "Developer",
              days: [8, 8, 4]
            }, {
              name: "James Smith",
              role: "Analyst",
              days: [4, 4, 4]
            }, {
              name: "Jim Smith",
              role: "Manager",
              days: [4, 4, 4]
            }
          ]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

        // extract date placeholder - interpolated into string referenced at B4
        expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString().toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Extracted on 41276");

        // revision placeholder - cell C4 changed from string to number
        expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text.toString()).toEqual("10");

        // dates placeholder - added cells
        expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text.toString()).toEqual("41275");
        expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text.toString()).toEqual("41276");
        expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text.toString()).toEqual("41277");

        // planData placeholder - added rows and cells
        expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John Smith");
        expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("James Smith");
        expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Jim Smith");

        expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Developer");
        expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Analyst");
        expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Manager");

        expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text.toString()).toEqual("4");

        expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text.toString()).toEqual("4");

        expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text.toString()).toEqual("4");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test1.xlsx', newData, 'binary');

        done();
      });

    });

    it("can substitute values with descendant properties and generate a file", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          demo: { extractDate: new Date("2013-01-02") },
          revision: 10,
          dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
          planData: [
            {
              name: "John Smith",
              role: { name: "Developer" },
              days: [8, 8, 4]
            }, {
              name: "James Smith",
              role: { name: "Analyst" },
              days: [4, 4, 4]
            }, {
              name: "Jim Smith",
              role: { name: "Manager" },
              days: [4, 4, 4]
            }
          ]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

        // extract date placeholder - interpolated into string referenced at B4
        expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Extracted on 41276");

        // revision placeholder - cell C4 changed from string to number
        expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text.toString()).toEqual("10");

        // dates placeholder - added cells
        expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text.toString()).toEqual("41275");
        expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text.toString()).toEqual("41276");
        expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text.toString()).toEqual("41277");

        // planData placeholder - added rows and cells
        expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John Smith");
        expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("James Smith");
        expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Jim Smith");

        expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Developer");
        expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Analyst");
        expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Manager");

        expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text.toString()).toEqual("4");

        expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text.toString()).toEqual("4");

        expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text.toString()).toEqual("4");
        expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text.toString()).toEqual("4");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test2.xlsx', newData, 'binary');

        done();
      });

    });

    it("can substitute values when single item array contains an object and generate a file", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't3.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          demo: { extractDate: new Date("2013-01-02") },
          revision: 10,
          planData: [
            {
              name: "John Smith",
              role: { name: "Developer" }
            }
          ]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:C7");

        // extract date placeholder - interpolated into string referenced at B4
        expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Extracted on 41276");

        // revision placeholder - cell C4 changed from string to number
        expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text.toString()).toEqual("10");

        // planData placeholder - added rows and cells
        expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John Smith");

        expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Developer");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test6.xlsx', newData, 'binary');

        done();
      });

    });

    it("can substitute values when single item array contains an object with sub array containing primatives and generate a file", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          demo: { extractDate: new Date("2013-01-02") },
          revision: 10,
          dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
          planData: [
            {
              name: "John Smith",
              role: { name: "Developer" },
              days: [8, 8, 4]
            }
          ]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F7");

        // extract date placeholder - interpolated into string referenced at B4
        expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Extracted on 41276");

        // revision placeholder - cell C4 changed from string to number
        expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text.toString()).toEqual("10");

        // dates placeholder - added cells
        expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text.toString()).toEqual("41275");
        expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text.toString()).toEqual("41276");
        expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text.toString()).toEqual("41277");

        // planData placeholder - added rows and cells
        expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John Smith");

        expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Developer");


        expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text.toString()).toEqual("8");
        expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text.toString()).toEqual("4");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test7.xlsx', newData, 'binary');

        done();
      });

    });

    it("moves columns left or right when filling lists", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 'test-cols.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          emptyCols: [],
          multiCols: ["one", "two"],
          singleCols: [10]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be set
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:E6");

        // C4 should have moved left, and the old B4 should now be deleted
        expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString()).toEqual("101");
        expect(sheet1.find("./sheetData/row/c[@r='C4']")).toBeNull();

        // C5 should have moved right, and the old B5 should now be expanded
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B5']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("one");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C5']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("two");
        expect(sheet1.find("./sheetData/row/c[@r='D5']/v").text.toString()).toEqual("102");

        // C6 should not have moved, and the old B6 should be replaced
        expect(sheet1.find("./sheetData/row/c[@r='B6']/v").text.toString()).toEqual("10");
        expect(sheet1.find("./sheetData/row/c[@r='C6']/v").text.toString()).toEqual("103");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test3.xlsx', newData, 'binary');

        done();
      });

    });

    it("moves rows down when filling tables", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 'test-tables.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute("Tables", {
          ages: [{ name: "John", age: 10 }, { name: "Bob", age: 2 }],
          scores: [{ name: "John", score: 100 }, { name: "Bob", score: 110 }, { name: "Jim", score: 120 }],
          coords: [],
          dates: [
            { name: "John", dates: [new Date("2013-01-01"), new Date("2013-01-02")] },
            { name: "Bob", dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")] },
            { name: "Jim", dates: [] },
          ]
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:H17");

        // Marker above table hasn't moved
        expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text.toString()).toEqual("101");

        // Headers on row 6 haven't moved
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B6']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Name");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='C6']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Age");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='E6']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Name");
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='F6']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Score");

        // Rows 7 contains table values for the two tables, plus the original marker in G7
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John");
        expect(sheet1.find("./sheetData/row/c[@r='C7']/v").text.toString()).toEqual("10");

        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='E7']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John");
        expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text.toString()).toEqual("100");

        expect(sheet1.find("./sheetData/row/c[@r='G7']/v").text.toString()).toEqual("102");

        // Row 8 contains table values, and no markers
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Bob");
        expect(sheet1.find("./sheetData/row/c[@r='C8']/v").text.toString()).toEqual("2");

        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='E8']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Bob");
        expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text.toString()).toEqual("110");

        expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

        // Row 9 contains no values for the first table, and again no markers
        expect(sheet1.find("./sheetData/row/c[@r='B9']")).toBeNull();
        expect(sheet1.find("./sheetData/row/c[@r='C9']")).toBeNull();

        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='E9']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Jim");
        expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text.toString()).toEqual("120");

        expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

        // Row 12 contains two blank cells and a marker
        expect(sheet1.find("./sheetData/row/c[@r='B12']/v")).toBeNull();
        expect(sheet1.find("./sheetData/row/c[@r='C12']/v")).toBeNull();
        expect(sheet1.find("./sheetData/row/c[@r='D12']/v").text.toString()).toEqual("103");

        // Row 15 contains a name, two dates, and a placeholder that was shifted to the right
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B15']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("John");
        expect(sheet1.find("./sheetData/row/c[@r='C15']/v").text.toString()).toEqual("41275");
        expect(sheet1.find("./sheetData/row/c[@r='D15']/v").text.toString()).toEqual("41276");
        expect(sheet1.find("./sheetData/row/c[@r='E15']/v").text.toString()).toEqual("104");

        // Row 16 contains a name and three dates
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B16']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Bob");
        expect(sheet1.find("./sheetData/row/c[@r='C16']/v").text.toString()).toEqual("41275");
        expect(sheet1.find("./sheetData/row/c[@r='D16']/v").text.toString()).toEqual("41276");
        expect(sheet1.find("./sheetData/row/c[@r='E16']/v").text.toString()).toEqual("41277");

        // Row 17 contains a name and no dates
        expect(
          sharedStrings.findall("./si")[
            parseInt(sheet1.find("./sheetData/row/c[@r='B17']/v").text.toString(), 10)
          ].find("t").text.toString()
        ).toEqual("Jim");
        expect(sheet1.find("./sheetData/row/c[@r='C17']")).toBeNull();

        // XXX: For debugging only
        fs.writeFileSync('test/output/test4.xlsx', newData, 'binary');

        done();
      });

    });

    it("replaces hyperlinks in sheet", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-hyperlinks.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          email: "john@bob.com",
          subject: "hello",
          url: "http://www.google.com",
          domain: "google"
        });

        const newData = t.generate(null);

        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        const rels = etree.parse(t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").asText()).getroot();

        // expect(sheet1.find("./hyperlinks/hyperlink/c[@r='C16']/v").text.toString()).toEqual("41275");
        expect(rels.find("./Relationship[@Id='rId2']").attrib.Target).toEqual("http://www.google.com");
        expect(rels.find("./Relationship[@Id='rId1']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");

        // XXX: For debugging only
        fs.writeFileSync('test/output/test9.xlsx', newData, 'binary');

        done();
      });
    });

    it("moves named tables, named cells and merged cells", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 'test-named-tables.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute("Tables", {
          ages: [
            { name: "John", age: 10 },
            { name: "Bill", age: 12 }
          ],
          days: ["Monday", "Tuesday", "Wednesday"],
          hours: [
            { name: "Bob", days: [10, 20, 30] },
            { name: "Jim", days: [12, 24, 36] }
          ],
          progress: 100
        });

        const newData = t.generate(null);

        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        const workbook = etree.parse(t.archive.file("xl/workbook.xml").asText()).getroot();
        const table1 = etree.parse(t.archive.file("xl/tables/table1.xml").asText()).getroot();
        const table2 = etree.parse(t.archive.file("xl/tables/table2.xml").asText()).getroot();
        const table3 = etree.parse(t.archive.file("xl/tables/table3.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:L29");

        // Named ranges have moved
        expect(workbook.find("./definedNames/definedName[@name='BelowTable']").text.toString()).toEqual("Tables!$B$18");
        expect(workbook.find("./definedNames/definedName[@name='Moving']").text.toString()).toEqual("Tables!$G$8");
        expect(workbook.find("./definedNames/definedName[@name='RangeBelowTable']").text.toString()).toEqual("Tables!$B$19:$C$19");
        expect(workbook.find("./definedNames/definedName[@name='RangeRightOfTable']").text.toString()).toEqual("Tables!$E$14:$F$14");
        expect(workbook.find("./definedNames/definedName[@name='RightOfTable']").text.toString()).toEqual("Tables!$F$8");

        // Merged cells have moved
        expect(sheet1.find("./mergeCells/mergeCell[@ref='B2:C2']")).not.toBeNull(); // title - unchanged

        expect(sheet1.find("./mergeCells/mergeCell[@ref='B10:C10']")).toBeNull(); // pushed down
        expect(sheet1.find("./mergeCells/mergeCell[@ref='B12:C12']")).not.toBeNull(); // pushed down

        expect(sheet1.find("./mergeCells/mergeCell[@ref='E7:F7']")).toBeNull(); // pushed down and accross
        expect(sheet1.find("./mergeCells/mergeCell[@ref='G8:H8']")).not.toBeNull(); // pushed down and accross

        // Table ranges and autofilter definitions have moved
        expect(table1.attrib.ref).toEqual("B4:C7"); // Grown
        expect(table1.find("./autoFilter").attrib.ref).toEqual("B4:C6"); // Grown

        expect(table2.attrib.ref).toEqual("B8:E10"); // Grown and pushed down
        expect(table2.find("./autoFilter").attrib.ref).toEqual("B8:E10"); // Grown and pushed down

        expect(table3.attrib.ref).toEqual("C14:D16"); // Grown and pushed down
        expect(table3.find("./autoFilter").attrib.ref).toEqual("C14:D16"); // Grown and pushed down

        // XXX: For debugging only
        fs.writeFileSync('test/output/test5.xlsx', newData, 'binary');

        done();
      });

    });

    it("Correctly parse when formula in the file", (done) => {

      fs.readFile(path.join(__dirname, 'templates', 'template.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);
        t.substitute(1, {
          people: [
            {
              name: "John Smith",
              age: 55,
            },
            {
              name: "John Doe",
              age: 35,
            }
          ]
        });

        done();
      });

    });

    it("Correctly recalculate formula", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-formula.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);
        t.substitute(1, {
          data: [
            { name: 'A', quantity: 10, unitCost: 3 },
            { name: 'B', quantity: 15, unitCost: 5 },
          ]
        });

        const newData = t.generate(null);
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        expect(sheet1).toBeDefined();

        expect(sheet1.find("./sheetData/row/c[@r='D2']/f").text.toString()).toEqual("Table3[Qty]*Table3[UnitCost]");
        expect(sheet1.find("./sheetData/row/c[@r='D2']/v")).toBeNull();

        // This part is not working
        // expect(sheet1.find("./sheetData/row/c[@r='D3']/f").text.toString()).toEqual("Table3[Qty]*Table3[UnitCost]");

        // fs.writeFileSync('test/output/test6.xlsx',newData, 'binary');
        done();
      });
    });

    it("File without dimensions works", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'gdocs.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);
        t.substitute(1, {
          planData: [
            { name: 'A', role: 'Role 1' },
            { name: 'B', role: 'Role 2' },
          ]
        });

        const newData = t.generate(null);
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        expect(sheet1).toBeDefined();

        // fs.writeFileSync('test/output/test7.xlsx',newData, 'binary');
        done();
      });
    });

    it("Array indexing", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-array.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);
        t.substitute(1, {
          data: [
            "First row",
            { name: 'B' },
          ]
        });

        const newData = t.generate(null);
        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        expect(sheet1).toBeDefined();
        expect(sheet1.find("./sheetData/row/c[@r='A2']/v")).not.toBeNull();
        expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("First row");
        expect(sheet1.find("./sheetData/row/c[@r='B2']/v")).not.toBeNull();
        expect(getSharedString(sharedStrings, sheet1, "B2")).toEqual("B");

        // fs.writeFileSync('test/output/test8.xlsx',newData, 'binary');
        done();
      });
    });

    it("Arrays with single element", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-nested-arrays.xlsx'), (err, buffer) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(buffer);
        const data = { "sales": [{ "payments": [123] }] };
        t.substitute(1, data);

        const newData = t.generate(null);
        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        expect(sheet1).toBeDefined();
        const a1 = sheet1.find("./sheetData/row/c[@r='A1']/v");
        const firstElement = sheet1.findall("./sheetData/row/c[@r='A1']");
        expect(a1).not.toBeNull();
        expect(a1.text.toString()).toEqual("123");
        expect(firstElement).not.toBeNull();
        expect(firstElement.length).toEqual(1);

        fs.writeFileSync('test/output/test-nested-arrays.xlsx', newData, 'binary');
        done();
      });
    });

    it("will correctly fill cells on all rows where arrays are used to dynamically render multiple cells", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), (err, data) => {
        expect(err).toBeNull();

        const t = new XlsxTemplate(data);

        t.substitute(1, {
          demo: { extractDate: new Date("2013-01-02") },
          revision: 10,
          dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
          planData: [{
            name: "John Smith",
            role: { name: "Developer" },
            days: [1, 2, 3]
          },
          {
            name: "James Smith",
            role: { name: "Analyst" },
            days: [1, 2, 3, 4, 5]
          },
          {
            name: "Jim Smith",
            role: { name: "Manager" },
            days: [1, 2, 3, 4, 5, 6, 7]
          }
          ]
        });

        const newData = t.generate(null);

        // var sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot(),
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();

        // Dimensions should be updated
        expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

        // Check length of all rows
        expect((sheet1.find("./sheetData/row[@r='7']") as any)._children.length).toEqual(2 + 3);
        expect((sheet1.find("./sheetData/row[@r='8']") as any)._children.length).toEqual(2 + 5);
        expect((sheet1.find("./sheetData/row[@r='9']") as any)._children.length).toEqual(2 + 7);

        fs.writeFileSync('test/output/test8.xlsx', newData, 'binary');

        done();
      });
    });

    it("do not move Images", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-move-images.xlsx'), (err, data) => {
        expect(err).toBeNull();
        const option = {
          moveImages: false
        }
        const t = new XlsxTemplate(data, option);
        t.substitute(1, {
          users: [
            {
              name: "John",
              surname: "Smith"
            },
            {
              name: "John",
              surname: "Doe"
            }
          ]
        });
        const newData = t.generate(null);
        const drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
        expect(drawingSheet).toBeDefined();
        drawingSheet.findall('xdr:twoCellAnchor').forEach(element => {
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("3");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("9");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("2");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("9");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("1");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("11");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("10");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("24");
          }
        });
        fs.writeFileSync('test/output/test_donotmoveImages.xlsx', newData, 'binary');
        done();
      });
    });

    it("Move Images", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-move-images.xlsx'), (err, data) => {
        expect(err).toBeNull();
        const option = {
          moveImages: true
        }
        const t = new XlsxTemplate(data, option);
        t.substitute(1, {
          users: [
            {
              name: "John",
              surname: "Smith"
            },
            {
              name: "John",
              surname: "Doe"
            }
          ]
        });
        const newData = t.generate(null);
        const drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
        expect(drawingSheet).toBeDefined();
        drawingSheet.findall('xdr:twoCellAnchor').forEach(element => {
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("4");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("10");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("2");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("9");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("1");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("11");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("11");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("25");
          }
        });
        fs.writeFileSync('test/output/test_moveImages.xlsx', newData, 'binary');
        done();
      });
    });

    it("Move Images with sameLine option", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-move-images.xlsx'), (err, data) => {
        expect(err).toBeNull();
        const option = {
          moveImages: true,
          moveSameLineImages: true,
        }
        const t = new XlsxTemplate(data, option);
        t.substitute(1, {
          users: [
            {
              name: "John",
              surname: "Smith"
            },
            {
              name: "John",
              surname: "Doe"
            }
          ]
        });
        const newData = t.generate(null);
        const drawingSheet = etree.parse(t.archive.file("xl/drawings/drawing1.xml").asText()).getroot();
        expect(drawingSheet).toBeDefined();
        drawingSheet.findall('xdr:twoCellAnchor').forEach(element => {
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='3']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("4");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("10");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='6']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("3");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("10");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='8']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("1");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("11");
          }
          if (element.find("xdr:pic/xdr:nvPicPr/xdr:cNvPr[@id='10']")) {
            expect(element.find("xdr:from/xdr:row").text.toString()).toEqual("11");
            expect(element.find("xdr:to/xdr:row").text.toString()).toEqual("25");
          }
        });
        fs.writeFileSync('test/output/test_moveImages_withSameLineOption.xlsx', newData, 'binary');
        done();
      });
    });

    it("Insert image and create rels", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-insert-images.xlsx'), (err, data) => {
        expect(err).toBeNull();
        const option = {
          imageRootPath: path.join(__dirname, 'templates', 'dataset')
        }
        const t = new XlsxTemplate(data, option);
        const imgB64 = 'iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC';

        t.substitute('init_rels', {
          imgB64,
        });
        const newData = t.generate(null);
        const rels = etree.parse(t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").asText()).getroot();
        expect(rels.findall("Relationship").length).toEqual(1);
        expect(rels.findall("Relationship")[0].attrib.Id).toEqual("rId1");
        expect(rels.findall("Relationship")[0].attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
        expect(rels.findall("Relationship")[0].attrib.Target).toEqual("../drawings/drawing2.xml");
        const drawing2 = etree.parse(t.archive.file("xl/drawings/drawing2.xml").asText()).getroot();
        expect(drawing2.findall("xdr:oneCellAnchor").length).toEqual(1);
        expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:from")[0].findall("xdr:col")[0].text.toString()).toEqual("1");
        expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:from")[0].findall("xdr:row")[0].text.toString()).toEqual("2");
        expect(drawing2.findall("xdr:oneCellAnchor")[0].findall("xdr:pic")[0].findall("xdr:blipFill")[0].findall("a:blip")[0].attrib["r:embed"]).toEqual("rId1");
        const relsdrawing2 = etree.parse(t.archive.file("xl/drawings/_rels/drawing2.xml.rels").asText()).getroot();
        expect(relsdrawing2.findall("Relationship").length).toEqual(1);
        expect(relsdrawing2.findall("Relationship")[0].attrib.Id).toEqual("rId1");
        expect(relsdrawing2.findall("Relationship")[0].attrib.Target).toEqual("../media/image1.jpg");
        expect(relsdrawing2.findall("Relationship")[0].attrib.Type).toEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
        // TODO : How can i compare the jpg file in the archive with my imgB64 variable ?
        // var image = t.archive.file("xl/media/image1.jpg");
        fs.writeFileSync('test/output/insert_image.xlsx', newData, 'binary');
        done();
      });
    });

    it("Insert some format of image", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-insert-images.xlsx'), (err, data) => {
        expect(err).toBeNull();
        const option = {
          imageRootPath: path.join(__dirname, 'templates', 'dataset')
        }
        const t = new XlsxTemplate(data, option);
        const imgB64 = 'iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC';

        t.substitute('multi_test', {
          imgB64,
          imgBuffer: Buffer.from(imgB64, 'base64'),
          imgPath: null, // "image1.png",
          high: null, // 'high.png',
          large: null, // "large.png",
          imgArray: [
            // { filename: "image1.png" },
            // { filename: "image2.png" },
            // { filename: "image3.png" },
            // { filename: "image4.png" },
          ],
          someText: "Hello Image",
        });
        const newData = t.generate(null);
        // TODO : make some test
        fs.writeFileSync('test/output/insert_images_format.xlsx', newData, 'binary');
        done();
      });
    });

    it("Insert images into table of merge cells", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-image-tables.xlsx'), (err, buffer) => {
        expect(err).toBeNull();
        const option = {
        }
        const t = new XlsxTemplate(buffer, option);
        const imgB64 = 'iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC';
        const data = {
          test: "Bug : If remove me, there are an error on ref.match([regex])",
          imgArray: []
        };
        for (let i = 0; i < 10; i++) {
          data.imgArray.push({ filename: imgB64 })
        }
        t.substitute('table_with_mergecell', data);
        const newData = t.generate(null);
        // TODO : make some test
        fs.writeFileSync('test/output/insert_image_table.xlsx', newData, 'binary');
        done();
      });
    });

    it("Insert 100 image", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'test-insert-images.xlsx'), (err, buffer) => {
        expect(err).toBeNull();
        const option = {
        }
        const t = new XlsxTemplate(buffer, option);
        const imgB64 = 'iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC';
        const data = {
          test: "Bug : If remove me, there are an error on ref.match([regex])",
          imgArray: []
        };
        for (let i = 0; i < 100; i++) {
          data.imgArray.push({ filename: imgB64 })
        }
        t.substitute('more_than_100', data);
        const newData = t.generate(null);

        const drawing2 = etree.parse(t.archive.file("xl/drawings/drawing2.xml").asText()).getroot();
        expect(drawing2.findall("xdr:oneCellAnchor").length).toEqual(100);
        const relsdrawing2 = etree.parse(t.archive.file("xl/drawings/_rels/drawing2.xml.rels").asText()).getroot();
        expect(relsdrawing2.findall("Relationship").length).toEqual(100);
        for (let i = 1; i < 101; i++) {
          const image = t.archive.file("xl/media/image" + i + ".jpg");
          expect(image).not.toBeFalsy();
        }
        fs.writeFileSync('test/output/insert_100_images.xlsx', newData, 'binary');
        done();
      });
    });
  });

  describe("Multiple sheets", () => {
    it("Each sheet should take each name", (done) => {
      fs.readFile(path.join(__dirname, 'templates', 'multple-sheets-arrays.xlsx'), (err, data) => {
        expect(err).toBeNull();

        // Create a template
        const t = new XlsxTemplate(data);
        for (let sheetNumber = 1; sheetNumber <= 2; sheetNumber++) {
          // Set up some placeholder values matching the placeholders in the template
          const values = {
            page: 'page: ' + sheetNumber,
            sheetNumber
          };

          // Perform substitution
          t.substitute(sheetNumber, values);
        }

        // Get binary data
        const newData = t.generate(null);
        const sharedStrings = etree.parse(t.archive.file("xl/sharedStrings.xml").asText()).getroot();
        const sheet1 = etree.parse(t.archive.file("xl/worksheets/sheet1.xml").asText()).getroot();
        const sheet2 = etree.parse(t.archive.file("xl/worksheets/sheet2.xml").asText()).getroot();
        expect(sheet1).toBeDefined();
        expect(sheet2).toBeDefined();
        expect(getSharedString(sharedStrings, sheet1, "A1")).toEqual("page: 1");
        expect(getSharedString(sharedStrings, sheet2, "A1")).toEqual("page: 2");
        expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("Page 1");
        expect(getSharedString(sharedStrings, sheet2, "A2")).toEqual("Page 2");

        fs.writeFileSync('test/output/multple-sheets-arrays.xlsx', newData, 'binary');
        done();
      });
    });
  });
});
