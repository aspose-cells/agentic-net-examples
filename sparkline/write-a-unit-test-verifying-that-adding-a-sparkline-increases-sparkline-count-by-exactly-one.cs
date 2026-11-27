// Title: C# unit test using Aspose.Cells to verify sparkline group count increases by one
// Common Searches: how to write a C# unit test for Aspose.Cells sparkline addition | Aspose.Cells verify sparkline collection size after adding a sparkline in C# | example of testing sparkline count increment with Aspose.Cells workbook
// Tags: Aspose.Cells sparkline collection count verification | C# Aspose.Cells add sparkline unit test | Aspose.Cells line sparkline group example | Aspose.Cells workbook sparkline unit testing | Aspose.Cells sparkline group manipulation API

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineUnitTest
{
    // The example creates a new Workbook, adds a line‑type sparkline group covering A1:D1, records the initial Sparklines.Count, adds a second sparkline at F1, and asserts that the group's Sparklines.Count has increased by exactly one, throwing an InvalidOperationException if the count is incorrect.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                SparklineTests.TestAddSparklineIncreasesCountByOne();
                Console.WriteLine("Test passed: Adding a sparkline increased the count by exactly one.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }

    public static class SparklineTests
    {
        public static void TestAddSparklineIncreasesCountByOne()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["C1"].PutValue(1);
            worksheet.Cells["D1"].PutValue(3);

            // Define the destination cell area where the sparkline will be placed (E1)
            CellArea destinationArea = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group of type Line
            int groupIndex = worksheet.SparklineGroups.Add(
                SparklineType.Line,          // Sparkline type
                "A1:D1",                     // Data range for the group (relative to the worksheet)
                false,                       // Show data in hidden rows/columns
                destinationArea);            // Destination area for the sparklines

            SparklineGroup group = worksheet.SparklineGroups[groupIndex];

            // Record the initial count of sparklines in the group (should be 1 because the destination area creates one sparkline)
            int initialCount = group.Sparklines.Count;

            // Add a new sparkline to a different cell (F1) to ensure the count increases
            // The Add method returns the index of the newly added sparkline
            group.Sparklines.Add("A1:D1", 0, 5); // Row 0, Column 5 (F1)

            // Record the count after adding
            int afterCount = group.Sparklines.Count;

            // Verify that the count increased by exactly one
            if (afterCount != initialCount + 1)
            {
                throw new InvalidOperationException(
                    $"Sparkline count mismatch. Expected {initialCount + 1}, but got {afterCount}.");
            }

            // Optional: Save the workbook to visually confirm (not required for the test)
            // workbook.Save("SparklineCountTest.xlsx", SaveFormat.Xlsx);
        }
    }
}
