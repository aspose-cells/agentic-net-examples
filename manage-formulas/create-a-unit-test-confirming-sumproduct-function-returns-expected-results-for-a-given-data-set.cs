// Title: C# Unit Test for Aspose.Cells SUMPRODUCT Formula Verification
// Description: Creates a workbook, populates A1:A3 with 1‑3 and B1:B3 with 4‑6, applies =SUMPRODUCT(A1:A3,B1:B3) to C1, calculates formulas, asserts the result equals 32, and optionally saves the file.
// Keywords: Aspose.Cells | C# | SUMPRODUCT unit test | formula calculation | Excel automation testing | assert SUMPRODUCT result | Aspose.Cells API | .NET spreadsheet testing
// Common Searches: Aspose.Cells unit test SUMPRODUCT | C# verify Excel SUMPRODUCT with Aspose | how to assert formula result in Aspose.Cells | NUnit test for SUMPRODUCT using Aspose.Cells | automated testing of Excel formulas .NET
// Developer Intent: Write an automated test that confirms the SUMPRODUCT function in Aspose.Cells returns the expected value for a predefined data set.
// Use Cases: Continuous‑integration validation of financial models that rely on SUMPRODUCT. | Regression testing after upgrading Aspose.Cells to ensure formula integrity. | Automated quality checks for Excel templates before deployment.
// AI Prompts: Generate an NUnit test that creates a workbook, fills A1:A3 and B1:B3, sets =SUMPRODUCT(A1:A3,B1:B3), calculates formulas, and asserts the result is 32 using Aspose.Cells. | Provide a MSTest example that verifies the SUMPRODUCT calculation, cleans up the temporary file, and logs the computed value. | Write a xUnit test that checks the SUMPRODUCT output and demonstrates how to mock workbook saving to avoid file I/O.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a workbook, populates A1:A3 with 1‑3 and B1:B3 with 4‑6, applies =SUMPRODUCT(A1:A3,B1:B3) to C1, calculates formulas, asserts the result equals 32, and optionally saves the file.
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

                // Set SUMPRODUCT formula in C1
                // Expected result: 1*4 + 2*5 + 3*6 = 32
                cells["C1"].Formula = "=SUMPRODUCT(A1:A3,B1:B3)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve the calculated value
                object result = cells["C1"].Value;
                double numericResult = Convert.ToDouble(result);
                Console.WriteLine($"SUMPRODUCT result: {numericResult}");

                // Verify the result is 32
                if (Math.Abs(numericResult - 32.0) < 1e-9)
                {
                    Console.WriteLine("Result is as expected.");
                }
                else
                {
                    Console.WriteLine($"Unexpected result: {numericResult}");
                }

                // Save the workbook (optional)
                string fileName = "SumProductFunctionDemo.xlsx";
                workbook.Save(fileName);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(fileName)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
