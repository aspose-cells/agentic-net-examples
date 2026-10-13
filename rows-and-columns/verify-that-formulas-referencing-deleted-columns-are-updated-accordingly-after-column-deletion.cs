// Title: How to automatically adjust Excel formulas after deleting a column using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that deletes a specific column in a workbook and verifies that all dependent formulas are updated by Aspose.Cells. | Show how to recalculate a worksheet after column removal and compare the original and updated formula strings in C#. | Create a C# example that asserts the formula reference changes from =C1*2 to =B1*2 and the resulting cell value after deleting column B with Aspose.Cells.
// Common Searches: Aspose.Cells C# update formula references when a column is removed | C# delete column in Excel workbook and keep formulas correct using Aspose.Cells | How to recalculate formulas after column deletion in Aspose.Cells for .NET | Verify formula adjustment after deleting column B in Aspose.Cells C# example | Aspose.Cells automatic formula shift after column removal .NET
// Tags: Aspose.Cells delete column with formula update | C# recalculate workbook after column removal | Excel formula reference auto‑adjust Aspose.Cells | verify updated formula value .NET | column deletion impact on dependent formulas

using System;
using Aspose.Cells;

// Demonstrates deleting column B in a workbook with Aspose.Cells for .NET, automatically updating the formula in D1 from =C1*2 to =B1*2, recalculating the sheet, and asserting the new value.
class FormulaUpdateAfterColumnDeletion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(5);   // Column A
            sheet.Cells["B1"].PutValue(10);  // Column B
            sheet.Cells["C1"].PutValue(15);  // Column C

            // Set a formula in D1 that references C1 (the third column)
            sheet.Cells["D1"].Formula = "=C1*2";

            // Calculate to evaluate the formula
            workbook.CalculateFormula();

            // Store original formula and value for later comparison
            string originalFormula = sheet.Cells["D1"].Formula;
            double originalValue = sheet.Cells["D1"].DoubleValue;

            // Delete column B (index 1, zero‑based). Columns to the right shift left.
            sheet.Cells.DeleteColumn(1);

            // Recalculate after column deletion so the updated formula is evaluated
            workbook.CalculateFormula();

            // After deletion, the original C column becomes B.
            // The formula should have been updated from "=C1*2" to "=B1*2".
            string updatedFormula = sheet.Cells["D1"].Formula;
            double updatedValue = sheet.Cells["D1"].DoubleValue;

            // Output verification results
            Console.WriteLine("Before deletion:");
            Console.WriteLine($"  Formula : {originalFormula}");
            Console.WriteLine($"  Value   : {originalValue}");

            Console.WriteLine("After deletion of column B:");
            Console.WriteLine($"  Updated Formula : {updatedFormula}");
            Console.WriteLine($"  Updated Value   : {updatedValue}");

            // Simple assertions (could be replaced with a testing framework)
            if (updatedFormula.Equals("=B1*2", StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(updatedValue - 30.0) < 0.0001)
            {
                Console.WriteLine("Formula updated correctly after column deletion.");
            }
            else
            {
                Console.WriteLine("Formula update after column deletion failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
