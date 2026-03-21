using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeImportDemo
{
    class Program
    {
        static void Main()
        {
            const string sourcePath = "SourceData.xlsx";

            // Ensure the source file exists; if not, create a sample workbook.
            if (!File.Exists(sourcePath))
            {
                var sampleWb = new Workbook();
                var sampleWs = sampleWb.Worksheets[0];
                var cells = sampleWs.Cells;

                // Fill sample data in A2:D10
                for (int r = 1; r <= 9; r++)          // rows 2‑10 (zero‑based index 1‑9)
                {
                    for (int c = 0; c < 4; c++)       // columns A‑D
                    {
                        cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                    }
                }

                sampleWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the source workbook (XLSX file)
            Workbook sourceWb = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sourceWs = sourceWb.Worksheets[0];

            // Define the range A2:D10 (zero‑based indices)
            int startRow = 1;      // Row 2
            int startColumn = 0;   // Column A
            int totalRows = 9;     // Rows 2‑10 inclusive
            int totalColumns = 4;  // Columns A‑D

            AsposeRange importRange = sourceWs.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Create a new workbook for the destination
            Workbook targetWb = new Workbook();
            Worksheet targetWs = targetWb.Worksheets[0];

            // Create a destination range starting at A1 with the same dimensions
            AsposeRange destRange = targetWs.Cells.CreateRange(0, 0, totalRows, totalColumns);

            // Copy values (including formulas) from the source range to the destination range
            destRange.Copy(importRange);

            // Save the target workbook
            targetWb.Save("RestrictedRegion.xlsx", SaveFormat.Xlsx);
        }
    }
}