// Title: Automatically adjust SUM formula range after deleting a row with Aspose.Cells for .NET
// AI Prompts: Delete the third row of a worksheet and let Aspose.Cells update any SUM formulas that reference the removed cells. | Recalculate the workbook and retrieve the updated formula text and value to confirm the range changed from A1:A5 to A1:A4. | Persist the workbook and output verification results indicating whether the formula reference and computed sum are correct.
// Common Searches: Aspose.Cells .NET how to keep SUM formula correct after removing a row | C# update cell references automatically when rows are deleted in an Excel workbook | Validate that formula range changes after DeleteRows in Aspose.Cells
// Tags: Aspose.Cells row deletion formula update | C# SUM range auto-adjustment | Excel recalculation after DeleteRows | check formula reference correctness | persist workbook changes

using Aspose.Cells;
using System;

// The example creates a workbook, fills column A with values 1‑5, sets B1 to =SUM(A1:A5), calculates the formula, deletes row 3, recalculates, verifies the formula updates to A1:A4 and the sum becomes 12, and finally saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with values 1 to 5
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Insert a formula in B1 that sums A1:A5
        sheet.Cells[0, 1].Formula = "SUM(A1:A5)";

        // Calculate the workbook to evaluate the formula
        workbook.CalculateFormula();

        // Store the original formula text and value
        string originalFormula = sheet.Cells[0, 1].Formula;
        double originalValue = sheet.Cells[0, 1].DoubleValue;

        // Delete row 3 (zero‑based index 2)
        sheet.Cells.DeleteRows(2, 1);

        // Re‑calculate after deletion
        workbook.CalculateFormula();

        // Get the updated formula and value
        string updatedFormula = sheet.Cells[0, 1].Formula;
        double updatedValue = sheet.Cells[0, 1].DoubleValue;

        // Verify that the formula reference has been adjusted (A1:A4)
        bool formulaAdjusted = updatedFormula.Contains("A4") && !updatedFormula.Contains("A5");

        // Verify that the value reflects the new sum (1+2+4+5 = 12)
        bool valueCorrect = Math.Abs(updatedValue - 12) < 0.0001;

        Console.WriteLine($"Original formula: {originalFormula}, value: {originalValue}");
        Console.WriteLine($"Updated formula: {updatedFormula}, value: {updatedValue}");
        Console.WriteLine($"Formula adjusted: {formulaAdjusted}");
        Console.WriteLine($"Value correct: {valueCorrect}");

        // Save the workbook (using the save rule)
        workbook.Save("Result.xlsx");
    }
}
