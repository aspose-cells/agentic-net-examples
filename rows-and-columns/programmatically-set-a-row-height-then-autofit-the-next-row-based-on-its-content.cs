using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – set row height then auto‑fit next row
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Row 0: set a custom height -----
        // Add some data to the first row (optional, just for demonstration)
        worksheet.Cells["A1"].PutValue("Custom height row");
        // Set the height of row 0 to 25 points (explicit height)
        worksheet.Cells.SetRowHeight(0, 25);

        // ----- Row 1: auto‑fit based on its content -----
        // Add longer text that will require a larger row height
        worksheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase when auto‑fitted.");
        // Auto‑fit row 1 (index 1) using the content of its cells
        worksheet.AutoFitRow(1);

        // Save the workbook (replace with your desired path)
        workbook.Save("RowHeightDemo.xlsx");
    }
}