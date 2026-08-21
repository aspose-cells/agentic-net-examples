// Title: C# Unit Test for SUMPRODUCT Formula Using Aspose.Cells
// Description: Demonstrates how to create a workbook with Aspose.Cells, populate ranges A1:A3 (1‑3) and B1:B3 (4‑6), apply the =SUMPRODUCT(A1:A3,B1:B3) formula to C1, calculate all formulas, assert that the result equals 32, and save the file as SumProductFunctionTest.xlsx.
// Keywords: Aspose.Cells | C# | .NET | SUMPRODUCT unit test | formula calculation | spreadsheet testing | Excel automation | assert formula result
// Common Searches: Aspose.Cells unit test SUMPRODUCT | C# verify SUMPRODUCT calculation | how to test Excel formulas with Aspose.Cells | assert SUMPRODUCT result in .NET | automated spreadsheet formula testing
// Developer Intent: Write an automated test that confirms the SUMPRODUCT function returns the expected value when evaluated by Aspose.Cells.
// Use Cases: Validate correctness of SUMPRODUCT after workbook calculation. | Include formula verification in continuous‑integration pipelines for spreadsheet‑related projects. | Detect regressions in custom Excel processing logic that relies on array formulas.
// AI Prompts: Generate an MSTest method in C# that builds a workbook, inserts =SUMPRODUCT(A1:A3,B1:B3), runs CalculateFormula, and asserts the result is 32. | Create an xUnit test using Aspose.Cells to populate two ranges, apply the SUMPRODUCT formula, calculate, and verify the output equals 32 before cleaning up the file. | Provide a NUnit example that sets up a workbook, adds the SUMPRODUCT formula, executes calculation, checks the value, and handles any exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates how to create a workbook with Aspose.Cells, populate ranges A1:A3 (1‑3) and B1:B3 (4‑6), apply the =SUMPRODUCT(A1:A3,B1:B3) formula to C1, calculate all formulas, assert that the result equals 32, and save the file as SumProductFunctionTest.xlsx.
    public class SumProductFunctionDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate first array: A1:A3 = {1, 2, 3}
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);
                cells["A3"].PutValue(3);

                // Populate second array: B1:B3 = {4, 5, 6}
                cells["B1"].PutValue(4);
                cells["B2"].PutValue(5);
                cells["B3"].PutValue(6);

                // Set SUMPRODUCT formula in C1 (expected result: 32)
                cells["C1"].Formula = "=SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve the calculated value
                object result = cells["C1"].Value;
                double numericResult = Convert.ToDouble(result);
                Console.WriteLine($"SUMPRODUCT result: {numericResult}");

                // Verify the result is 32
                if (Math.Abs(numericResult - 32.0) > 0.0001)
                {
                    Console.WriteLine("Verification failed: result is not 32.");
                }
                else
                {
                    Console.WriteLine("Verification succeeded: result is 32.");
                }

                // Save the workbook
                string fileName = "SumProductFunctionTest.xlsx";
                workbook.Save(fileName);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(fileName)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
