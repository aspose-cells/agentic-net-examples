// Title: Aspose.Cells for .NET – Copy Rows 5‑10 to 20‑25 with Formatting (C#)
// Description: C# example that creates a workbook, fills rows 5‑10 with values and bold styling, then uses Aspose.Cells CopyRows to duplicate those rows into rows 20‑25 while preserving all cell data, styles, and merged cells. The workbook is saved as RowCopyResult.xlsx.
// Keywords: Aspose.Cells copy rows C# | CopyRows method .NET | preserve cell formatting Aspose | duplicate worksheet rows | Aspose.Cells example GitHub | C# Excel row copy
// Common Searches: Aspose.Cells copy rows with formatting C# | How to duplicate rows 5 to 10 to 20 to 25 Aspose.Cells | CopyRows preserve styles .NET | C# copy Excel rows using Aspose.Cells | Aspose.Cells copy rows example GitHub
// Developer Intent: Duplicate rows 5‑10 into rows 20‑25 while keeping all cell values and styles intact.
// Use Cases: Replicate a formatted template block for repeated report sections. | Create a snapshot of specific rows before applying calculations. | Insert a pre‑styled data segment elsewhere in the sheet without losing formatting.
// AI Prompts: Write C# code with Aspose.Cells to copy rows 5‑10 to 20‑25, keep all formatting, then remove the original rows. | Show how to copy rows that contain merged cells and conditional formatting using Aspose.Cells CopyRows. | Provide an example of copying rows between different worksheets while preserving styles in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowCopyDemo
{
    // C# example that creates a workbook, fills rows 5‑10 with values and bold styling, then uses Aspose.Cells CopyRows to duplicate those rows into rows 20‑25 while preserving all cell data, styles, and merged cells. The workbook is saved as RowCopyResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate rows 5 to 10 (zero‑based indices 4‑9) with sample data and formatting
            for (int row = 4; row <= 9; row++)
            {
                // Sample values
                cells[row, 0].PutValue($"Row {row + 1} Col A");
                cells[row, 1].PutValue($"Row {row + 1} Col B");

                // Sample formatting: bold font for the first column
                Style style = workbook.CreateStyle();
                style.Font.IsBold = true;
                cells[row, 0].SetStyle(style);
            }

            // Define source and destination parameters
            int sourceRowIndex = 4;          // Row 5 (zero‑based)
            int destinationRowIndex = 19;   // Row 20 (zero‑based)
            int rowCount = 6;               // Number of rows to copy (rows 5‑10)

            // Copy rows preserving data and formatting
            cells.CopyRows(cells, sourceRowIndex, destinationRowIndex, rowCount);

            // Save the workbook
            workbook.Save("RowCopyResult.xlsx");
        }
    }
}
