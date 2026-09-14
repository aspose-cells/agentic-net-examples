// Title: C# unit test to verify that adding a sparkline to an Aspose.Cells worksheet increments the sparkline collection count by one
// AI Prompts: Generate a C# unit test (MSTest/NUnit) that creates a workbook with Aspose.Cells, adds a sparkline group, records the initial SparklineCollection count, adds a sparkline, and asserts the count increased by exactly one. | Write code for a test method that validates the SparklineCollection.Add method updates the collection size correctly in Aspose.Cells.
// Common Searches: aspocells unit test sparkline collection count increment | c# verify sparkline added count Aspose.Cells | how to assert sparkline count after adding sparkline in .NET | unit testing Aspose.Cells sparkline group in MSTest
// Tags: aspocells sparklinecollection add test | c# aspocells sparkline count verification | aspocells sparklinegroup unit testing | mstest sparkline collection size assertion | dotnet sparkline addition unit test

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates a C# unit test that creates a workbook with Aspose.Cells, fills sample data, adds a sparkline group, records the initial SparklineCollection count, adds a sparkline, and asserts that the count increased by exactly one, throwing an exception if the verification fails.
    public class SparklineTests
    {
        public static void Main()
        {
            try
            {
                AddingSparkline_IncreasesCountByOne();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public static void AddingSparkline_IncreasesCountByOne()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (row 0, columns A‑E)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[0, i].PutValue(i + 1);
            }

            // Define where the sparkline will be placed (cell A2)
            CellArea location = new CellArea
            {
                StartRow = 1,
                EndRow = 1,
                StartColumn = 0,
                EndColumn = 0
            };

            // Add a sparkline group that references the data range A1:E1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];
            SparklineCollection sparklines = group.Sparklines;

            // Record the initial count (should be 0)
            int initialCount = sparklines.Count;

            // Add a sparkline to the group
            sparklines.Add("A1:E1", 1, 0);

            // Verify that the count increased by exactly one
            int afterCount = sparklines.Count;
            if (afterCount != initialCount + 1)
            {
                throw new InvalidOperationException($"Expected count {initialCount + 1}, but got {afterCount}.");
            }
        }
    }
}
