using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header row
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");

        // Data rows (some rows/columns will be hidden)
        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Alice");
        cells["C2"].PutValue(85);

        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Bob");
        cells["C3"].PutValue(92);

        cells["A4"].PutValue(3);
        cells["B4"].PutValue("Charlie");
        cells["C4"].PutValue(78);

        // Hide row 3 (index 2) and column B (index 1)
        sheet.Cells.HideRow(2);      // hides the row containing Bob
        sheet.Cells.HideColumn(1);   // hides the "Name" column

        // Determine the area that contains visible data.
        CellArea exportArea = new CellArea
        {
            StartRow = 0,
            EndRow = sheet.Cells.MaxDataRow,
            StartColumn = 0,
            EndColumn = sheet.Cells.MaxDataColumn
        };

        // Configure JsonSaveOptions
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportArea = exportArea,   // limit export to the used range
            SkipEmptyRows = true,      // skip rows that are completely empty
            ExportEmptyCells = false,  // do not include empty cells
            HasHeaderRow = true        // first row is a header
        };

        // Save the workbook as JSON
        string jsonPath = "ExportedData.json";
        workbook.Save(jsonPath, jsonOptions);
    }
}