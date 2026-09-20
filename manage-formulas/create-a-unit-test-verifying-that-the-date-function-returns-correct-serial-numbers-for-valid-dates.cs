// Title: Write a C# unit test with Aspose.Cells that validates the Excel DATE function returns the correct serial number
// AI Prompts: Generate a C# NUnit test that creates a Workbook, sets cell A1 formula to =DATE(2023,5,10) using Aspose.Cells, calls CalculateFormula, and asserts the cell value equals the expected OADate double within a tolerance. | Produce a C# xUnit test method that verifies Aspose.Cells correctly converts a given year, month, and day to Excel's serial date by comparing the calculated result with DateTime.ToOADate.
// Common Searches: aspocells c# unit test for excel date function serial number | how to assert DATE formula result with Aspose.Cells in a test | c# verify Excel DATE returns correct OADate using Aspose.Cells CalculateFormula | unit testing Excel date serial conversion with Aspose.Cells library
// Tags: Aspose.Cells DATE formula unit test | C# Excel OADate validation with Aspose | CalculateFormula date serial verification | unit test workbook formula evaluation | Excel serial number comparison C#

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a new Workbook, writes a DATE formula for a specific year, month, and day into cell A1, runs CalculateFormula, retrieves the resulting double value, and asserts that it matches the .NET OADate serial number within a small tolerance, providing a clear pass/fail outcome.
    public class DateFunctionDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Define a date to test
                int year = 2023;
                int month = 5;
                int day = 10;

                // Expected serial number using .NET OADate (same as Excel's serial)
                double expectedSerial = new DateTime(year, month, day).ToOADate();

                // Set the DATE formula in cell A1
                worksheet.Cells["A1"].Formula = $"=DATE({year},{month},{day})";

                // Calculate the formula
                workbook.CalculateFormula();

                // Retrieve the calculated value
                var actualValue = worksheet.Cells["A1"].Value;

                // Validate result
                if (actualValue is double actualSerial)
                {
                    const double tolerance = 0.000001;
                    if (Math.Abs(expectedSerial - actualSerial) <= tolerance)
                    {
                        Console.WriteLine($"Test passed. Serial number: {actualSerial}");
                    }
                    else
                    {
                        Console.WriteLine($"Test failed. Expected: {expectedSerial}, Actual: {actualSerial}");
                    }
                }
                else
                {
                    Console.WriteLine("Test failed. Result is not a double.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
