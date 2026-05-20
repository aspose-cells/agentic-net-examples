using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineMoveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the original location of the sparkline (cell E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the original location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // At this point the group already contains one sparkline at E1.
            // Retrieve its index (first sparkline)
            int sparklineIdx = 0; // only one sparkline exists

            // ----- Move the sparkline -----
            // 1. Remove the existing sparkline
            group.Sparklines.RemoveAt(sparklineIdx);

            // 2. Define the new location (cell G1)
            int newRow = 0;      // same row
            int newColumn = 6;   // Column G (0‑based index)

            // 3. Re‑add the sparkline at the new location
            group.Sparklines.Add("A1:D1", newRow, newColumn);

            // Optional: verify the new sparkline position
            Sparkline movedSparkline = group.Sparklines[0];
            Console.WriteLine($"Sparkline moved to Row: {movedSparkline.Row}, Column: {movedSparkline.Column}");

            // Save the workbook
            workbook.Save("SparklineMoved.xlsx");
        }
    }
}