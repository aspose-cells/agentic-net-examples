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

            // Define the location where the original sparkline will be placed (E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the data range A1:D1 and the original location
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // The group already contains a sparkline at the original location (E1)
            Sparkline originalSparkline = group.Sparklines[0];

            // ------------------------------------------------------------
            // Copy the sparkline to a new location (E3) – row index 2, column 4
            // ------------------------------------------------------------
            // Sparkline.Row and Sparkline.Column are read‑only, so we add a new sparkline
            // with the same data range at the desired location.
            int copiedSparklineIndex = group.Sparklines.Add(originalSparkline.DataRange, 2, 4);
            Sparkline copiedSparkline = group.Sparklines[copiedSparklineIndex];

            // Optional: remove the original sparkline if you only need the copy
            // group.Sparklines.RemoveAt(0);

            // Output positions for verification
            Console.WriteLine($"Original sparkline location: Row={originalSparkline.Row}, Column={originalSparkline.Column}");
            Console.WriteLine($"Copied sparkline location:   Row={copiedSparkline.Row}, Column={copiedSparkline.Column}");

            // Save the workbook
            workbook.Save("SparklineCopyWithLocation.xlsx");
        }
    }
}