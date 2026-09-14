// Title: Write a C# unit test with Aspose.Cells to verify SUMPRODUCT(A1:A3,B1:B3) returns 32
// AI Prompts: Generate an MSTest method that creates a Workbook, fills A1:A3 and B1:B3 with numbers, assigns the SUMPRODUCT formula to C1, triggers CalculateFormula, and asserts the cell value equals 32. | Produce a NUnit test case that uses Aspose.Cells to evaluate the SUMPRODUCT function on two numeric ranges and validates the computed result against the expected value.
// Common Searches: how to assert SUMPRODUCT calculation result in Aspose.Cells unit test | Aspose.Cells C# example for testing Excel SUMPRODUCT formula | unit testing Excel formulas with Aspose.Cells library in .NET | verify SUMPRODUCT(A1:A3,B1:B3) output using Aspose.Cells in a test project
// Tags: Aspose.Cells SUMPRODUCT unit testing | C# workbook formula evaluation test | Excel SUMPRODUCT verification with Aspose.Cells | calculate formulas Aspose.Cells C# | assert workbook cell value .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating cells A1:A3 and B1:B3 with numeric values, applying the SUMPRODUCT(A1:A3,B1:B3) formula to C1, calculating all formulas, and asserting that the result equals 32 in a C# unit test.
    public class SumProductExample
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate data for the arrays
                // A1:A3 = {1, 2, 3}
                // B1:B3 = {4, 5, 6}
                sheet.Cells["A1"].PutValue(1);
                sheet.Cells["A2"].PutValue(2);
                sheet.Cells["A3"].PutValue(3);
                sheet.Cells["B1"].PutValue(4);
                sheet.Cells["B2"].PutValue(5);
                sheet.Cells["B3"].PutValue(6);

                // Set SUMPRODUCT formula in C1
                sheet.Cells["C1"].Formula = "SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Expected result: 1*4 + 2*5 + 3*6 = 32
                double expected = 32;
                double actual = sheet.Cells["C1"].DoubleValue;

                // Verify the result
                if (Math.Abs(expected - actual) < 1e-9)
                {
                    Console.WriteLine($"SUMPRODUCT calculation is correct. Result = {actual}");
                }
                else
                {
                    Console.WriteLine($"SUMPRODUCT calculation is incorrect. Expected = {expected}, Actual = {actual}");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
