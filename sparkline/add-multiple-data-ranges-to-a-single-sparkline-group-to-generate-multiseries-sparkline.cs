using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate first data range (Series 1) ----------
            // Range: A1:E5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue(row + col + 1);
                }
            }

            // ---------- Populate second data range (Series 2) ----------
            // Range: F1:J5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 5; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col - 4));
                }
            }

            // Define the location range for the multi‑series sparkline.
            // It must have the same number of rows as the data ranges (5 rows).
            CellArea location = new CellArea
            {
                StartRow = 5,   // Row 6 (zero‑based)
                EndRow = 9,     // Row 10 (zero‑based) → 5 rows total
                StartColumn = 0,
                EndColumn = 0   // Single column (A)
            };

            // Add a sparkline group with the first data range.
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E5", false, location);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Include the second data range to create a multi‑series sparkline.
            sparklineGroup.ResetRanges("A1:E5,F1:J5", false, location);

            // Save the workbook containing the multi‑series sparkline.
            string outputPath = "MultiSeriesSparkline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}