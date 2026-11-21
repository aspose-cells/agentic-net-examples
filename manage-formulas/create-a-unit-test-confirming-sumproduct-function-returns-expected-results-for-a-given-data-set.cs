// Title: C# unit test that validates Aspose.Cells SUMPRODUCT formula returns 66 for a two‑column array
// Common Searches: how to write a C# unit test for SUMPRODUCT using Aspose.Cells | Aspose.Cells calculate SUMPRODUCT formula in a workbook example | verify Excel SUMPRODUCT result with Aspose.Cells .NET API | unit testing array multiplication formula in Aspose.Cells | C# code to assert SUMPRODUCT calculation with Aspose.Cells
// Tags: Aspose.Cells SUMPRODUCT formula validation | C# unit test for Excel array multiplication | workbook.CalculateFormula with SUMPRODUCT | Aspose.Cells formula evaluation example | verify numeric array product in .NET workbook

using System;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a new Workbook, populates cells A1:A3 and B1:B3 with numbers, sets the SUMPRODUCT(A1:A3,B1:B3) formula in C1, triggers workbook.CalculateFormula, retrieves the computed value, and asserts that it equals the expected result of 66.0, demonstrating how to unit‑test Excel formula calculations with Aspose.Cells for .NET.
    public class SumProductDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate first array (A1:A3)
                sheet.Cells["A1"].PutValue(2);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(5);

                // Populate second array (B1:B3)
                sheet.Cells["B1"].PutValue(4);
                sheet.Cells["B2"].PutValue(6);
                sheet.Cells["B3"].PutValue(8);

                // Set SUMPRODUCT formula in C1
                sheet.Cells["C1"].Formula = "SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve the calculated value
                double result = sheet.Cells["C1"].DoubleValue;
                const double expected = 66.0;

                // Verify the result
                if (Math.Abs(result - expected) < 0.0001)
                {
                    Console.WriteLine($"Success: Result = {result}");
                }
                else
                {
                    Console.WriteLine($"Failure: Expected {expected}, but got {result}");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
