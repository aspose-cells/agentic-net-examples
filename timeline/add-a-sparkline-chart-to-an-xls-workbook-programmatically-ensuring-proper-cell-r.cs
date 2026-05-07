using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:J10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Define the location range where sparklines will be placed (column K, rows 1-10)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                EndRow = 9,     // Row 10
                StartColumn = 10, // Column K (zero‑based)
                EndColumn = 10
            };

            // Add a sparkline group: line type, data range A1:J10, horizontal orientation (isVertical = false)
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:J10",
                false,
                sparklineLocation);

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Optional: customize the sparkline group (show high/low points, set colors, line weight)
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            group.HighPointColor.Color = System.Drawing.Color.Green;
            group.LowPointColor.Color = System.Drawing.Color.Red;
            group.LineWeight = 1.0;

            // Save the workbook; format is inferred from the file extension
            workbook.Save("SparklineExample.xls");
        }
    }
}