// Title: C# Unit Test – Verify SparklineCollection Count Increments by One in Aspose.Cells
// Description: Creates a workbook, adds a line‑type SparklineGroup, records the initial SparklineCollection count, adds a sparkline, and asserts the count increased by exactly one, throwing an error otherwise.
// Keywords: Aspose.Cells sparkline unit test | SparklineCollection Add verification | C# Aspose.Cells sparkline count | unit test for sparkline addition | .NET spreadsheet sparkline testing
// Common Searches: how to test sparkline count increase Aspose.Cells C# | unit test SparklineCollection.Add Aspose.Cells | verify sparkline collection size after add .NET | Aspose.Cells sparkline unit testing example
// Developer Intent: Confirm that adding a sparkline to a SparklineCollection raises its Count property by exactly one.
// Use Cases: Integrate into CI pipelines to catch off‑by‑one regressions in sparkline handling. | Validate custom report generators that programmatically add sparklines. | Ensure library upgrades do not break sparkline collection behavior.
// AI Prompts: Generate an MSTest/NUnit/xUnit test that asserts SparklineCollection.Count increments after adding a sparkline with Aspose.Cells. | Refactor the sample to use Assert.AreEqual (or equivalent) instead of manual exception handling. | Explain how to mock Workbook and Worksheet for faster sparkline unit tests.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // Creates a workbook, adds a line‑type SparklineGroup, records the initial SparklineCollection count, adds a sparkline, and asserts the count increased by exactly one, throwing an error otherwise.
    public class SparklineCollectionTests
    {
        public void AddingSparkline_IncreasesCountByOne()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline (A1:D1)
                worksheet.Cells["A1"].PutValue(5);
                worksheet.Cells["B1"].PutValue(2);
                worksheet.Cells["C1"].PutValue(1);
                worksheet.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed (cell E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4, // Column E (0‑based index)
                    EndColumn = 4
                };

                // Add a sparkline group of type Line with the data range A1:D1
                int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup group = worksheet.SparklineGroups[groupIndex];

                // Get the sparkline collection from the group
                SparklineCollection sparklines = group.Sparklines;

                // Record the initial count (should be 0)
                int initialCount = sparklines.Count;

                // Add a sparkline to the collection at E1
                sparklines.Add("A1:D1", 0, 4);

                // Verify that the count increased by exactly one
                if (sparklines.Count != initialCount + 1)
                {
                    throw new InvalidOperationException($"Sparkline count expected {initialCount + 1} but was {sparklines.Count}.");
                }

                // Optional: Save the workbook for manual verification
                // workbook.Save("SparklineAdditionTest.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during test: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                var test = new SparklineCollectionTests();
                test.AddingSparkline_IncreasesCountByOne();
                Console.WriteLine("Test passed: Adding sparkline increased count by one.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }
}
