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
