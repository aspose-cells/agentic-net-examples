// Title: C# Aspose.Cells – Verify SUMPRODUCT support before using the formula
// Description: Creates a workbook, fills A1:A3 and B1:B3, assigns =SUMPRODUCT(A1:A3,B1:B3) to C1, uses the HasCustomFunction property to detect if SUMPRODUCT is supported, calculates the result when possible, and saves the file.
// Keywords: Aspose.Cells SUMPRODUCT support | HasCustomFunction C# | detect unsupported Excel functions .NET | validate Excel function compatibility | Aspose.Cells formula checking | C# Excel function detection | Aspose.Cells example GitHub | Excel SUMPRODUCT Aspose.Cells
// Common Searches: how to check if SUMPRODUCT is supported in Aspose.Cells | Aspose.Cells HasCustomFunction usage | detect custom functions in Aspose.Cells C# | verify Excel function compatibility with Aspose.Cells | C# example for unsupported Excel functions
// Developer Intent: Identify whether the SUMPRODUCT function is available in the current Aspose.Cells version before performing calculations.
// Use Cases: Prevent runtime errors by confirming SUMPRODUCT support prior to large‑scale data processing. | Implement a fallback algorithm when SUMPRODUCT is reported as unsupported. | Log or report any custom (unsupported) functions encountered while importing workbooks.
// AI Prompts: Write C# code that checks any Excel formula for unsupported functions using HasCustomFunction and provides an alternative calculation. | Generate a method that scans all worksheets in a workbook and returns formulas containing custom functions. | Create an Aspose.Cells example that replaces unsupported SUMPRODUCT formulas with equivalent supported expressions.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills A1:A3 and B1:B3, assigns =SUMPRODUCT(A1:A3,B1:B3) to C1, uses the HasCustomFunction property to detect if SUMPRODUCT is supported, calculates the result when possible, and saves the file.
    class CheckSumProductSupport
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the SUMPRODUCT calculation
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(4);
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(6);

            // Set a formula that uses SUMPRODUCT
            Cell formulaCell = cells["C1"];
            formulaCell.Formula = "=SUMPRODUCT(A1:A3,B1:B3)";

            // Check whether the function is recognized as a custom (unsupported) function
            if (formulaCell.HasCustomFunction)
            {
                Console.WriteLine("SUMPRODUCT is NOT supported by the current Aspose.Cells version.");
            }
            else
            {
                Console.WriteLine("SUMPRODUCT is supported. Calculating...");

                // Calculate the workbook formulas
                workbook.CalculateFormula();

                // Output the result
                Console.WriteLine("Result of SUMPRODUCT(A1:A3,B1:B3): " + formulaCell.Value);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CheckSumProductSupport.xlsx");
        }
    }
}
