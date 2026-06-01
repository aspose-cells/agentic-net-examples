using System;
using Aspose.Cells;

class CopyRowsAndAutoFitDemo
{
    static void Main()
    {
        // Create source workbook and populate it with sample data
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Sample data that will affect row height
        sourceSheet.Cells["A1"].PutValue("This is a long text that should increase the row height significantly.");
        sourceSheet.Cells["A2"].PutValue("Short");
        sourceSheet.Cells["A3"].PutValue("Multi-line\ntext\nexample");

        // Set explicit row heights to illustrate that they will be copied
        sourceSheet.Cells.Rows[0].Height = 30; // Row 0
        sourceSheet.Cells.Rows[1].Height = 15; // Row 1
        sourceSheet.Cells.Rows[2].Height = 40; // Row 2

        // Create destination workbook
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Determine how many rows to copy (rows that contain data)
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;

        // Copy rows from source to destination (starting at row 0 in both sheets)
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, rowsToCopy);

        // Auto‑fit the copied rows in the destination sheet so their heights match the content
        destinationSheet.AutoFitRows(0, rowsToCopy - 1);

        // Save the resulting workbook
        destinationWorkbook.Save("CopiedAndAutoFitRows.xlsx");
    }
}