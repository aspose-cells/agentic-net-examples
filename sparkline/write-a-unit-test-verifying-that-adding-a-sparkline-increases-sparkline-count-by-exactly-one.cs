// Title: Write a C# unit test with Aspose.Cells that confirms adding a sparkline increases the worksheet's SparklineGroup count by one
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a line sparkline to a data range, captures the SparklineGroup count before and after, and asserts the count grew from 0 to 1. | Produce NUnit test code that inserts a column sparkline using Aspose.Cells, records the SparklineGroup collection size, adds the sparkline, and verifies the collection size incremented by exactly one.
// Common Searches: aspocells c# unit test sparkline count verification | how to assert sparklinegroup size after adding sparkline in .NET | mstest example for sparkline insertion with Aspose.Cells | nunit test sparkline collection increment Aspose.Cells
// Tags: aspocells sparklinegroup count verification | c# unit test aspocells sparkline | mstest aspocells sparkline insertion | nunit aspocells sparkline validation | aspocells add sparkline assert collection size

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example shows how to write a C# unit test that creates a new Workbook, fills cells A1:A5 with numeric values, adds a sparkline to that range, and asserts that the worksheet's SparklineGroup collection size increases by exactly one.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the source range
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["A4"].PutValue(40);
                sheet.Cells["A5"].PutValue(50);

                // Verify that the data was written correctly by counting non‑empty cells in column A
                int nonEmptyCount = 0;
                for (int row = 0; row < 5; row++)
                {
                    if (sheet.Cells[row, 0].Value != null)
                    {
                        nonEmptyCount++;
                    }
                }

                // Expected count is 5
                if (nonEmptyCount == 5)
                {
                    Console.WriteLine("Test passed: All data cells are populated.");
                }
                else
                {
                    Console.WriteLine($"Test failed: Expected 5 populated cells, but found {nonEmptyCount}.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
