// Title: Ignore #REF! Errors with Aspose.Cells CalculationOptions in C#
// Description: Demonstrates how to suppress #REF! formula errors in a workbook using Aspose.Cells. The example creates a workbook, adds a value to A1, sets a formula in B1, deletes column A to generate a #REF! reference, configures CalculationOptions.IgnoreError = true, runs workbook.CalculateFormula, and reads the resulting cell value.
// Keywords: Aspose.Cells | CalculationOptions | IgnoreError | #REF! error handling | C# formula evaluation | .NET spreadsheet API | suppress formula errors | delete column reference | workbook.CalculateFormula | Excel error handling
// Common Searches: Aspose.Cells ignore #REF! error | CalculationOptions IgnoreError C# example | how to suppress formula errors Aspose.Cells | calculate workbook with invalid references .NET | Aspose.Cells delete column and recalculate
// Developer Intent: Calculate all formulas in an Aspose.Cells workbook while ignoring #REF! and other formula errors.
// Use Cases: Processing templates where some references may be removed but the workbook must still be calculated. | Batch converting or analyzing spreadsheets without interruption from broken formulas. | Generating reports that retain error strings (e.g., "#REF!") instead of throwing exceptions.
// AI Prompts: Show C# code that uses Aspose.Cells CalculationOptions.IgnoreError to evaluate formulas containing #REF! references. | Explain what happens when IgnoreError is true and a formula results in a #REF! error in Aspose.Cells. | Provide a step‑by‑step example of deleting a column, triggering a #REF! error, and then calculating the workbook without raising an exception.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationOptionsDemo
{
    // Demonstrates how to suppress #REF! formula errors in a workbook using Aspose.Cells. The example creates a workbook, adds a value to A1, sets a formula in B1, deletes column A to generate a #REF! reference, configures CalculationOptions.IgnoreError = true, runs workbook.CalculateFormula, and reads the resulting cell value.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a value in A1
            cells["A1"].PutValue(10);

            // Set a formula in B1 that references A1
            cells["B1"].Formula = "=A1*2";

            // Delete column A to make the reference in B1 invalid (#REF!)
            sheet.Cells.DeleteColumn(0);

            // At this point B1 contains a formula that refers to a deleted column,
            // which results in a #REF! error when calculated.

            // Create calculation options and set IgnoreError to true
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true   // Suppress errors during calculation
            };

            // Calculate all formulas in the workbook using the options
            workbook.CalculateFormula(calcOptions);

            // Output the result of the formula in B1.
            // With IgnoreError = true, the engine will ignore the #REF! error
            // and keep the cell's value as the error string "#REF!".
            Console.WriteLine("Result of B1 after ignoring errors: " + cells["B1"].StringValue);

            // (Optional) Save the workbook to verify the result in Excel
            workbook.Save("CalculationOptionsIgnoreErrorDemo.xlsx");
        }
    }
}
