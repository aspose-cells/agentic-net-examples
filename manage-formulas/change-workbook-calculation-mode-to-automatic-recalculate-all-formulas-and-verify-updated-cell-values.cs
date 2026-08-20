// Title: Aspose.Cells .NET – Switch Workbook to Automatic Calculation, Recalculate All Formulas, and Verify Cell Values
// Description: Demonstrates how to change a workbook's calculation mode to Automatic using Aspose.Cells for .NET, trigger a full recalculation with CalculateFormula, read the updated cell values, and save the workbook.
// Keywords: Aspose.Cells automatic calculation | Workbook.CalculateFormula .NET | set calculation mode Aspose.Cells | recalculate formulas C# | verify formula results Aspose.Cells | Aspose.Cells workbook settings | C# Excel formula evaluation
// Common Searches: Aspose.Cells change calculation mode to automatic | How to recalculate all formulas in Aspose.Cells | Read updated cell values after CalculateFormula | Aspose.Cells verify formula output programmatically | C# set workbook calculation mode Aspose.Cells
// Developer Intent: The developer needs to enable automatic calculation for a workbook, force a full formula recompute, and confirm that the resulting cell values are correct before saving.
// Use Cases: Ensure all dependent formulas are up‑to‑date when exporting a workbook. | Programmatically recalculate after modifying input cells in a data‑processing pipeline. | Validate formula results in automated tests or CI builds.
// AI Prompts: Generate C# code that sets the workbook calculation mode to Manual, updates a cell, then forces a recalculation with Aspose.Cells. | Provide a unit‑test method that asserts expected values of formula cells after calling workbook.CalculateFormula. | Explain how to capture and handle calculation errors when using workbook.CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to change a workbook's calculation mode to Automatic using Aspose.Cells for .NET, trigger a full recalculation with CalculateFormula, read the updated cell values, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data and formulas
        cells["A1"].PutValue(5);               // Simple value
        cells["B1"].Formula = "=A1*2";         // Depends on A1
        cells["C1"].Formula = "=B1+10";        // Depends on B1

        // Change calculation mode to Automatic (FormulaSettings.CalculationMode)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Verify that the cell values have been updated
        Console.WriteLine("A1 value: " + cells["A1"].Value); // Expected: 5
        Console.WriteLine("B1 value: " + cells["B1"].Value); // Expected: 10
        Console.WriteLine("C1 value: " + cells["C1"].Value); // Expected: 20

        // Save the workbook (lifecycle save rule)
        workbook.Save("UpdatedWorkbook.xlsx");
    }
}
