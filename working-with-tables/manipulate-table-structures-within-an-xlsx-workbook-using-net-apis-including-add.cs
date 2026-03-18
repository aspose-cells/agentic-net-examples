using System;
using Aspose.Cells;

class TableManipulationExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate initial 3x3 data table
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["C1"].PutValue("Header3");
        cells["A2"].PutValue(1);
        cells["B2"].PutValue(2);
        cells["C2"].PutValue(3);
        cells["A3"].PutValue(4);
        cells["B3"].PutValue(5);
        cells["C3"].PutValue(6);
        cells["A4"].PutValue(7);
        cells["B4"].PutValue(8);
        cells["C4"].PutValue(9);

        // Insert a new row at index 2 (third row) – existing rows shift down
        cells.InsertRow(2);
        // Fill the inserted row with data
        cells["A3"].PutValue(100);
        cells["B3"].PutValue(200);
        cells["C3"].PutValue(300);

        // Delete the original fourth row (now at index 4 after insertion)
        cells.DeleteRow(4);

        // Insert two new columns after column B (index 1)
        cells.InsertColumns(2, 2);
        // Set headers for the new columns
        cells["D1"].PutValue("NewCol1");
        cells["E1"].PutValue("NewCol2");
        // Populate the new columns with sample data
        for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
        {
            cells[row, 3].PutValue(row * 10); // Column D
            cells[row, 4].PutValue(row * 20); // Column E
        }

        // Delete the first column (index 0) – all columns shift left
        cells.DeleteColumn(0);

        // Modify a specific cell value
        cells["B2"].PutValue("Modified");

        // Save the workbook (lifecycle: save)
        workbook.Save("TableManipulationResult.xlsx");
    }
}