// Title: How to test that Aspose.Cells manual calculation mode stops automatic formula evaluation until Workbook.CalculateFormula is called (C#)
// AI Prompts: Create a C# Aspose.Cells workbook, set CalcModeType.Manual, add a formula, and print the cell value before and after invoking Workbook.CalculateFormula. | Show how to disable automatic formula calculation in Aspose.Cells and then explicitly trigger evaluation with Workbook.CalculateFormula, verifying the computed result.
// Common Searches: Aspose.Cells C# manual calculation mode example showing formula not evaluated automatically | disable automatic formula calculation in Aspose.Cells until CalculateFormula is called | C# Aspose.Cells workbook.CalculateFormula after setting CalcModeType.Manual | testing manual calc mode behavior with formulas in Aspose.Cells .NET | how to verify manual calculation mode prevents formula evaluation in Aspose.Cells
// Tags: manual calculation mode Aspose.Cells | Workbook.CalculateFormula C# example | disable automatic formula evaluation Aspose.Cells | set CalcModeType.Manual .NET | formula evaluation timing Aspose.Cells

using System;
using Aspose.Cells;

namespace ManualCalcModeTest
{
    // The example creates a workbook, places values in A1 and A2, assigns a formula to B1, switches the workbook to manual calculation mode, demonstrates that B1 remains empty before calling Workbook.CalculateFormula, then evaluates the formula to produce 15, prints the before/after values, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(10);

            // Set a formula that depends on the data above
            cells["B1"].Formula = "=A1+A2";

            // Set calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // At this point the formula should NOT be evaluated automatically.
            // The cell's value will be null (or default) because we haven't called CalculateFormula.
            Console.WriteLine("Before CalculateFormula:");
            Console.WriteLine($"B1 value (expected empty): {(cells["B1"].Value ?? "null")}");

            // Explicitly calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Now the formula should be evaluated and the result should be 15.
            Console.WriteLine("After CalculateFormula:");
            Console.WriteLine($"B1 value (expected 15): {cells["B1"].IntValue}");

            // Save the workbook (optional, just to demonstrate saving)
            workbook.Save("ManualCalcModeTest.xlsx");
        }
    }
}
