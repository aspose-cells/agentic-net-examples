// Title: C# unit test that validates the SUMPRODUCT formula result using Aspose.Cells
// Common Searches: how to unit test SUMPRODUCT formula with Aspose.Cells in C# | Aspose.Cells calculate SUMPRODUCT in a .NET unit test | verify Excel SUMPRODUCT result programmatically using Aspose.Cells | C# workbook.CalculateFormula example for array multiplication | testing Excel formula evaluation with Aspose.Cells unit tests
// Tags: SUMPRODUCT formula verification Aspose.Cells | Aspose.Cells calculate formula unit test | C# workbook.CalculateFormula SUMPRODUCT | Excel array multiplication test Aspose.Cells | Aspose.Cells formula result assertion .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills A1:A3 with 1‑3 and B1:B3 with 4‑6, inserts the SUMPRODUCT(A1:A3,B1:B3) formula in C1, runs workbook.CalculateFormula, and asserts that the computed value equals 32, reporting success or failure.
    public class SumProductDemo
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

                // Insert SUMPRODUCT formula in C1
                sheet.Cells["C1"].Formula = "SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Expected result: 1*4 + 2*5 + 3*6 = 32
                double expected = 32;
                double actual = sheet.Cells["C1"].DoubleValue;

                // Verify that the calculated value matches the expected result
                if (Math.Abs(expected - actual) < 1e-9)
                {
                    Console.WriteLine("SUMPRODUCT calculation succeeded. Result = " + actual);
                }
                else
                {
                    Console.WriteLine($"SUMPRODUCT calculation failed. Expected = {expected}, Actual = {actual}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during execution: " + ex.Message);
            }
        }
    }
}
