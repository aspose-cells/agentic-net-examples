using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["C1"].PutValue(1);
            worksheet.Cells["D1"].PutValue(3);

            // Define where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line
            int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = worksheet.SparklineGroups[groupIndex];

            // Capture the initial sparkline count (should be 0)
            int initialCount = group.Sparklines.Count;

            // Add a sparkline to the group
            group.Sparklines.Add("A1:D1", 0, 4);

            // Capture the count after addition
            int afterCount = group.Sparklines.Count;

            // Verify that the count increased by exactly one
            if (afterCount == initialCount + 1)
            {
                Console.WriteLine("Test passed: Sparkline count increased by one.");
            }
            else
            {
                Console.WriteLine($"Test failed: Expected count {initialCount + 1}, but got {afterCount}.");
            }
        }
    }
}