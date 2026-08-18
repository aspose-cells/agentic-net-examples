// Title: Aspose.Cells C# – Verify Formula Text Remains Unchanged While Forcing Cell Value to Zero
// Description: Shows how to use Aspose.Cells' SetFormula overload to assign "=SUM(1,2)" to A1, set its cached result to 0, confirm the Formula property is unchanged, and then recalculate the workbook to reveal the true result before saving.
// Keywords: Aspose.Cells | C# | .NET | SetFormula overload | custom cell result | force cell value zero | preserve formula text | formula verification | Workbook.CalculateFormula | Excel automation
// Common Searches: Aspose.Cells set formula and custom result | keep formula string unchanged after SetFormula | force cell value to zero in Aspose.Cells | C# verify formula text after manual value assignment | calculate workbook after overriding cell result
// Developer Intent: Confirm that a cell’s formula string stays intact while its cached value is manually set to zero.
// Use Cases: Display a placeholder (e.g., zero) for a formula‑driven cell without altering the underlying expression. | Create test workbooks where formulas are preserved but results are overridden for validation pipelines. | Prepare spreadsheets with dummy results before performing a full calculation to speed up downstream processing.
// AI Prompts: Generate C# code that sets a formula with Aspose.Cells, forces the cell value to zero, verifies the formula text, then recalculates to obtain the actual result. | Explain how Aspose.Cells' SetFormula(string, object) overload can be used to assign a custom result while keeping the original formula unchanged. | Provide a step‑by‑step guide for asserting that a cell’s Formula property is unchanged after manually setting its Value in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    // Shows how to use Aspose.Cells' SetFormula overload to assign "=SUM(1,2)" to A1, set its cached result to 0, confirm the Formula property is unchanged, and then recalculate the workbook to reveal the true result before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a formula
            string formula = "=SUM(1,2)";

            // Set the formula and explicitly set its calculated result to zero
            // This uses the SetFormula(string, object) overload.
            cells["A1"].SetFormula(formula, 0);

            // Verify that the formula text has not changed
            bool formulaUnchanged = cells["A1"].Formula == formula;

            // Verify that the cell's value is zero (as we set it)
            bool valueIsZero = cells["A1"].Value is double d && d == 0.0;

            Console.WriteLine($"Formula unchanged: {formulaUnchanged}");
            Console.WriteLine($"Value is zero: {valueIsZero}");

            // Optionally calculate the workbook to see the real result after calculation
            workbook.CalculateFormula();

            // After calculation the value should reflect the actual formula result (3)
            Console.WriteLine($"Value after calculation: {cells["A1"].Value}");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FormulaVerification.xlsx");
        }
    }
}
