using System;
using Aspose.Cells;

class FreezeAfterAutoFit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate column width changes
        sheet.Cells["A1"].PutValue("Short");
        sheet.Cells["B1"].PutValue("A much longer text that will cause column B to expand significantly");
        sheet.Cells["C1"].PutValue("Medium length text");

        // Auto‑fit all columns so their widths match the content
        sheet.AutoFitColumns();

        // Freeze panes after autofit to lock the column widths in place.
        // Freeze at cell B2 (row index 1, column index 1) with 1 frozen row and 1 frozen column.
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook
        workbook.Save("FreezeAfterAutoFit.xlsx");
    }
}