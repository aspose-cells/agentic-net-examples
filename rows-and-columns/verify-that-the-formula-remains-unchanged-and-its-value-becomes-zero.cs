// Title: Aspose.Cells .NET – Preserve Formula while Forcing Displayed Value to Zero
// Description: Demonstrates how to keep a cell's formula unchanged (e.g., =SUM(1,2,3)) and assign a temporary displayed value of zero using the SetFormula(string, object) overload. The example verifies the formula string, checks the pre‑calculation value, runs workbook.CalculateFormula() to reveal the original result, and saves the workbook.
// Keywords: Aspose.Cells SetFormula overload | C# keep formula unchanged | force zero value cell | displayed value override | Aspose.Cells .NET example | cell formula preservation | temporary cell value
// Common Searches: Aspose.Cells set cell value without changing formula | SetFormula overload zero value .NET | keep formula unchanged and set displayed value | override cell value temporarily Aspose.Cells | how to force zero in a formula cell
// Developer Intent: Assign a zero displayed value to a cell while leaving its original formula intact.
// Use Cases: Show a placeholder zero in a formula cell before the workbook is calculated. | Create templates where formulas are retained but initial values are forced to zero for reporting or UI purposes. | Write unit tests that need to mock a cell's value without altering the underlying calculation logic.
// AI Prompts: Write C# code that uses Aspose.Cells SetFormula overload to set a custom value while preserving the original formula. | Explain how to confirm that a cell's formula string remains unchanged after assigning a zero value with SetFormula. | Show the steps to temporarily set a cell's displayed value to zero, then calculate the workbook to retrieve the actual formula result.

using System;
using Aspose.Cells;

// Demonstrates how to keep a cell's formula unchanged (e.g., =SUM(1,2,3)) and assign a temporary displayed value of zero using the SetFormula(string, object) overload. The example verifies the formula string, checks the pre‑calculation value, runs workbook.CalculateFormula() to reveal the original result, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1
        Cell cell = worksheet.Cells["A1"];

        // Set a normal formula (e.g., sum of constants)
        cell.Formula = "=SUM(1,2,3)"; // Expected result is 6

        // Keep the formula unchanged but set its displayed/calculated value to zero
        // This uses the SetFormula(string formula, object value) overload
        cell.SetFormula(cell.Formula, 0);

        // Verify that the formula string is still the same
        Console.WriteLine("Formula after SetFormula: " + cell.Formula);

        // Verify that the cell's value is zero before any calculation
        Console.WriteLine("Value before calculation: " + cell.Value);

        // Optionally calculate the workbook to see the actual result of the formula
        workbook.CalculateFormula();

        // Verify the value after calculation (should be 6, showing that the formula was unchanged)
        Console.WriteLine("Value after calculation: " + cell.Value);

        // Save the workbook (lifecycle rule)
        workbook.Save("FormulaZeroDemo.xlsx");
    }
}
