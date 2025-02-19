# XLSX Template for Browsers

> This is a fork from [xlsx-template by Optilude](https://github.com/optilude/xlsx-template), a Excel template replace library for Node.js. This fork has fixed some issues and supports being run in a browser.

This module provides a means of generating "real" Excel reports (i.e. not CSV
files) in NodeJS applications.

The basic principle is this: You create a template in Excel. This can be
formatted as you wish, contain formulae etc. In this file, you put placeholders
using a specific syntax (see below). In code, you build a map of placeholders
to values and then load the template, substitute the placeholders for the
relevant values, and generate a new .xlsx file that you can then serve to the
user.

## Placeholders

Placeholders are inserted in cells in a spreadsheet. It does not matter how
those cells are formatted, so e.g. it is OK to insert a placeholder (which is
text content) into a cell formatted as a number or currecy or date, if you
expect the placeholder to resolve to a number or currency or date.

### Scalars

Simple placholders take the format `${name}`. Here, `name` is the name of a
key in the placeholders map. The value of this placholder here should be a
scalar, i.e. not an array or object. The placeholder may appear on its own in a
cell, or as part of a text string. For example:

    | Extracted on: | ${extractDate} |

might result in (depending on date formatting in the second cell):

    | Extracted on: | Jun-01-2013 |

Here, `extractDate` may be a date and the second cell may be formatted as a
number.

Inside scalars there possibility to use array indexers. 
For example: 

Given data

    var template = { extractDates: ["Jun-01-2113", "Jun-01-2013" ]}

which will be applied to following template

    | Extracted on: | ${extractDates[0]} |

will results in the 

    | Extracted on: | Jun-01-2113 |

### Columns

You can use arrays as placeholder values to indicate that the placeholder cell
is to be replicated across columns. In this case, the placeholder cannot appear
inside a text string - it must be the only thing in its cell. For example,
if the placehodler value `dates` is an array of dates:

    | ${dates} |

might result in:

    | Jun-01-2013 | Jun-02-2013 | Jun-03-2013 |

### Tables

Finally, you can build tables made up of multiple rows. In this case, each
placeholder should be prefixed by `table:` and contain both the name of the
placeholder variable (a list of objects) and a key (in each object in the list).
For example:

    | Name                 | Age                 |
    | ${table:people.name} | ${table:people.age} |

If the replacement value under `people` is an array of objects, and each of
those objects have keys `name` and `age`, you may end up with something like:

    | Name        | Age |
    | John Smith  | 20  |
    | Bob Johnson | 22  |

If a particular value is an array, then it will be repeated across columns as
above.

### Images

You can insert images with   

    | My image: | ${image:imageData} |

Given data

    var template = { imageData: "<base64String>"}

You can insert a (vertical) list of images with   

    | My images | ${table:images.data:image} |

Given data

    var template = { images: [{ data : "<base64String>" }, { data : "<base64String>" }]}

Image data can be provided as: 
- Base64 string
- Base64 Buffer

If the image placeholder is in a standard cell, then the image is insert with its original size (may break the cell boundaries).  
If the image placeholder is in a merge cell, then the image is resized (without changing the aspect ratio) to fill the merge cell (similar to the CSS statement `object-fit: contain`).

You can pass an option `imageRatio` to resize the original image dimensions (in percent). This has no effect on merge cells.
 
    var option = {imageRatio : 75.4}  
    ...  
    var t = new XlsxTemplate(data, option);


## Generating reports

To make this magic happen, you need some code like this:

```
    var XlsxTemplate = require('xlsx-template');

    // Load an XLSX file into memory
    fs.readFile(path.join(__dirname, 'templates', 'template1.xlsx'), function(err, data) {

        // Create a template
        var template = new XlsxTemplate(data);

        // Replacements take place on first sheet
        var sheetNumber = 1;

        // Set up some placeholder values matching the placeholders in the template
        var values = {
                extractDate: new Date(),
                dates: [ new Date("2013-06-01"), new Date("2013-06-02"), new Date("2013-06-03") ],
                people: [
                    {name: "John Smith", age: 20},
                    {name: "Bob Johnson", age: 22}
                ]
            };

        // Perform substitution
        template.substitute(sheetNumber, values);

        // Get binary data
        var data = template.generate();

        // ...

    });
```

At this stage, `data` is a string blob representing the compressed archive that
is the `.xlsx` file (that's right, a `.xlsx` file is a zip file of XML files,
if you didn't know). You can send this back to a client, store it to disk,
attach it to an email or do whatever you want with it.

You can pass options to `generate()` to set a different return type. use
`{type: 'uint8array'}` to generate a `Uint8Array`, `arraybuffer`, `blob`,
`nodebuffer` to generate an `ArrayBuffer`, `Blob` or `nodebuffer`, or
`base64` to generate a base64-encoded string.

## Caveats

* The spreadsheet must be saved in `.xlsx` format. `.xls`, `.xlsb` or `.xlsm`
  won't work.
* Column (array) and table (array-of-objects) insertions cause rows and cells to
  be inserted or removed. When this happens, only a limited number of
  adjustments are made:
    * Merged cells and named cells/ranges to the right of cells where insertions
      or deletions are made are moved right or left, appropriately. This may
      not work well if cells are merged across rows, unless all rows have the
      same number of insertions.
    * Merged cells, named tables or named cells/ranges below rows where further
      rows are inserted are moved down.
  Formulae are not adjusted.
* As a corollary to this, it is not always easy to build formulae that refer
  to cells in a table (e.g. summing all rows) where the exact number of rows
  or columns is not known in advance. There are two strategies for dealing
  with this:
    * Put the table as the last (or only) thing on a particular sheet, and
      use a formula that includes a large number of rows or columns in the
      hope that the actual table will be smaller than this number.
    * Use named tables. When a placeholder in a named table causes columns or
      rows to be added, the table definition (i.e. the cells included in the
      table) will be updated accordingly. You can then use things like
      `TableName[ColumnName]` in your formula to refer to all values in a given
      column in the table as a logical range.
* Placeholders only work in simple cells and tables, pivot tables or
  other such things.

## Changelog

### Version 2.0.0
* Converted project to TypeScript
* Changed project target from Node.js to Browser
* Fixed minor issues regarding images

### Version 1.3.0
* Added support for optional moving of the images together with table. (#109)

### Version 1.2.0
* Specify license field in addition to licenses field in the package.json (#102)

### Version 1.1.0
* Added TypeScript definitions. #101
* NodeJS 12, 14 support

### Version 1.0.0
Nothing to see here. Just I'm being brave and make version 1.0.0

### Version 0.5.0
* Placeholder in hyperlinks. #87
* NodeJS 10 support

### Version 0.4.0
* Fix wrongly replacing text in shared strings #81

### Version 0.2.0

* Add ability copy and delete sheets.

### Version 0.0.7

* Fix bug with calculating <dimensions /> when adding columns

### Version 0.0.6

* You can now pass `options` to `generate()`, which are passed to JSZip
* Fix setting of sheet <dimensions /> when growing the sheet
* Fix corruption of sheet when writing dates
* Fix corruption of sheet when calculating calcChain

### Version 0.0.5

* Mysterious

### Version 0.0.4

Merged pending pull requests

* Deletion of the sheets.

### Version 0.0.3

Merged a number of overdue pull requests, including:

* Windows support
* Support for table footers
* Documentation improvements

### Version 0.0.2

* Fix a potential issue with the typing of string indices that could cause the
  first string to not render correctly if it contained a substitution.

### Version 0.0.1

* Initial release
