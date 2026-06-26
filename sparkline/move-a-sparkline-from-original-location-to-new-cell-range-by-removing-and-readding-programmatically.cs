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

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the initial location range for the sparkline (cell E1)
            CellArea initialLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the data range and initial location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, initialLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group at the initial location (row 0, column 4)
            int sparklineIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparklineIdx];

            Console.WriteLine($"Original sparkline position - Row: {sparkline.Row}, Column: {sparkline.Column}");

            // ----- Move the sparkline -----
            // Define the new location (cell E3 -> row 2, column 4)
            int newRow = 2;
            int newColumn = 4;

            // Remove the existing sparkline from the collection
            group.Sparklines.RemoveSparkline(sparkline);

            // Add a new sparkline at the desired location using the same data range
            int newSparklineIdx = group.Sparklines.Add("A1:D1", newRow, newColumn);
            Sparkline movedSparkline = group.Sparklines[newSparklineIdx];

            Console.WriteLine($"Moved sparkline position - Row: {movedSparkline.Row}, Column: {movedSparkline.Column}");

            // Save the workbook with the moved sparkline
            workbook.Save("SparklineMoved.xlsx");
        }
    }
}