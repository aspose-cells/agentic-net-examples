# XML Maps Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

XML Maps


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **XML Maps**.

Example:

create-a-workbook.cs


## Required Namespaces

Most examples will require:

using Aspose.Cells;


## Common Pattern

Typical Aspose.Cells workflow:

Workbook workbook = new Workbook();

Worksheet sheet = workbook.Worksheets[0];

Cells cells = sheet.Cells;


## Output

Examples may generate:

- XLSX files
- PDF files
- CSV files
- Images

Output files are written to the working directory.
- import-an-xml-map-from-an-xsd-file-into-a-workbook-using-workbookimportxml-method.cs
- import-an-xml-map-directly-from-an-xml-file-into-a-workbook-with-workbookimportxml.cs
- list-all-xml-maps-in-the-workbook-and-output-each-maps-name-to-the-console.cs
- retrieve-the-root-element-name-of-a-specific-map-by-its-index-using-workbookxmlmapsindexrootelementname.cs
- query-cells-mapped-to-a-given-xpath-expression-using-worksheetxmlmapquery-method.cs
- query-cells-with-namespaceaware-xpath-by-providing-prefix-mappings-to-worksheetxmlmapquery.cs
- validate-linked-cells-after-importing-xml-data-by-reexecuting-worksheetxmlmapquery-and-checking-results.cs
- export-xml-data-for-a-specific-map-to-a-file-using-workbookexportxml-with-map-index.cs
- export-xml-data-for-a-specific-map-to-a-memory-stream-using-workbookexportxml-overload.cs
- loop-through-all-xml-maps-in-a-workbook-and-export-each-maps-xml-using-exportxml.cs
- export-xml-with-utf8-encoding-by-specifying-encodingutf8-parameter-in-exportxml-method.cs
- export-xml-with-indentation-enabled-by-configuring-exportxmloptions-before-calling-exportxml.cs
- remove-an-unwanted-xml-map-from-the-workbook-by-its-index-using-xmlmapsremoveat.cs
- remove-an-xml-map-by-its-name-using-a-helper-that-searches-the-xmlmaps-collection.cs
- update-an-existing-xml-map-by-readding-it-with-the-same-name-and-a-new-xsd-schema.cs
- load-a-workbook-from-a-memory-stream-add-an-xml-map-and-save-back-to-a-stream.cs
- batch-process-a-folder-of-workbooks-adding-the-same-xml-map-to-each-file-programmatically.cs
- batch-export-xml-data-from-multiple-workbooks-by-iterating-files-and-invoking-exportxml-for-each-map.cs
- import-xml-data-into-linked-cells-using-workbookimportxml-after-the-xml-map-has-been-added.cs
- import-xml-data-from-a-stream-into-a-workbook-with-linked-cells-using-importxml-overload.cs
- export-xml-without-xml-declaration-by-setting-exportxmloptionsomitxmldeclaration-to-true.cs
- enumerate-the-xmlmapcollection-and-log-each-maps-name-and-root-element-name-for-debugging.cs
- load-an-excel-workbook-from-a-xlsx-file-and-attach-an-xml-schema-map.cs
- create-a-new-workbook-add-a-worksheet-and-define-an-xml-map-using-a-xsd-file.cs
