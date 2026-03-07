using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location range for the original sparkline (cell E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the original location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the original sparkline (row 0, column 4 -> E1)
            int sparkIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline originalSpark = group.Sparklines[sparkIdx];

            Console.WriteLine($"Original sparkline placed at Row={originalSpark.Row}, Column={originalSpark.Column}");

            // ----- Copy the sparkline to a new target cell (e.g., G3) -----
            // Target cell coordinates (row 2, column 6) -> G3 (0‑based indices)
            int targetRow = 2;    // third row (index 2)
            int targetColumn = 6; // column G (index 6)

            // Add a new sparkline with the same data range at the target location
            int copiedSparkIdx = group.Sparklines.Add(originalSpark.DataRange, targetRow, targetColumn);
            Sparkline copiedSpark = group.Sparklines[copiedSparkIdx];

            Console.WriteLine($"Copied sparkline placed at Row={copiedSpark.Row}, Column={copiedSpark.Column}");

            // Save the workbook
            workbook.Save("SparklineCopyDemo.xlsx");
        }
    }
}