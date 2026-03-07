using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineLocationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the cell area where the sparkline will be placed (initially empty)
            CellArea locationArea = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group with the data range A1:D1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locationArea);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Specify the exact location of the sparkline using row and column indices
            // Row = 2 (third row, zero‑based), Column = 5 (sixth column, zero‑based)
            int sparklineIdx = group.Sparklines.Add("A1:D1", 2, 5);
            Sparkline sparkline = group.Sparklines[sparklineIdx];

            // Output the location to verify
            Console.WriteLine($"Sparkline placed at Row: {sparkline.Row}, Column: {sparkline.Column}");

            // Save the workbook as XLSX
            workbook.Save("SparklineLocationDemo.xlsx");
        }
    }
}