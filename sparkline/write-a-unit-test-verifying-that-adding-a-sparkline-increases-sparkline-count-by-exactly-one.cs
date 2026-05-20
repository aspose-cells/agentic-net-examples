using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    public class SparklineTests
    {
        public void AddingSparkline_IncreasesCountByOne()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(3);
                sheet.Cells["D1"].PutValue(4);

                // Define the cell where the sparkline will be placed (E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group to the worksheet
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Get the sparkline collection from the group
                SparklineCollection sparklines = group.Sparklines;

                // Record the initial count (should be zero)
                int initialCount = sparklines.Count;

                // Add a sparkline to the collection
                sparklines.Add("A1:D1", 0, 4);

                // Verify that the count increased by exactly one
                if (sparklines.Count != initialCount + 1)
                {
                    throw new InvalidOperationException(
                        $"Expected sparkline count {initialCount + 1}, but got {sparklines.Count}.");
                }

                Console.WriteLine("Test passed: Sparkline count increased by one.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during test execution: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            var test = new SparklineTests();
            test.AddingSparkline_IncreasesCountByOne();
        }
    }
}