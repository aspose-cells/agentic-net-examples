using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location cell for the original sparkline (E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group with the data range and the original location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the original sparkline at row 0, column 4 (cell E1)
            int sparklineIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline originalSparkline = group.Sparklines[sparklineIdx];

            Console.WriteLine($"Original sparkline placed at Row={originalSparkline.Row}, Column={originalSparkline.Column}");

            // ----- Copy the sparkline to a new destination cell -----
            // Define the new destination cell (G1) -> row 0, column 6
            int destRow = 0;
            int destColumn = 6;

            // Add a new sparkline with the same data range but at the new location
            int copiedSparklineIdx = group.Sparklines.Add(originalSparkline.DataRange, destRow, destColumn);
            Sparkline copiedSparkline = group.Sparklines[copiedSparklineIdx];

            Console.WriteLine($"Copied sparkline placed at Row={copiedSparkline.Row}, Column={copiedSparkline.Column}");

            // Save the workbook to an XLSX file
            workbook.Save("SparklineCopyDemo.xlsx");
        }
    }
}