// Title: Aspose.Cells for .NET – Manual Calculation Mode with Forced Full Recalculation
// Description: Demonstrates how to switch a workbook to manual calculation (CalcModeType.Manual), modify cell values, enable ForceFullCalculation, and invoke workbook.CalculateFormula() to recompute all dependent formulas in a single pass. Includes performance tips for large spreadsheets and saving the result.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | CalcModeType.Manual | ForceFullCalculation | CalculateFormula | workbook.CalculateFormula | disable automatic calculation | performance optimization | large spreadsheet recalculation
// Common Searches: Aspose.Cells set manual calculation mode | ForceFullCalculation example C# | How to recalculate workbook after data changes Aspose.Cells | CalculateFormula manual mode .NET | Improve performance by disabling auto calculation Aspose.Cells
// Developer Intent: I need to turn off automatic formula evaluation, change cell data, and then recalculate the entire workbook on demand.
// Use Cases: Batch‑update thousands of cells in a financial model, then run a single full calculation to obtain final results. | Generate a report where intermediate formulas are postponed until all input data is populated, ensuring consistent output. | Optimize performance of large Excel files by using manual mode during data import and forcing a final recalculation before saving.
// AI Prompts: Show C# code that sets Aspose.Cells workbook to manual calculation, updates cells, enables ForceFullCalculation, and calls CalculateFormula. | Explain why manual calculation mode improves performance in Aspose.Cells and how to trigger a full recomputation. | Provide a step‑by‑step guide to batch modify data in a workbook and then force a complete recalculation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to switch a workbook to manual calculation (CalcModeType.Manual), modify cell values, enable ForceFullCalculation, and invoke workbook.CalculateFormula() to recompute all dependent formulas in a single pass. Includes performance tips for large spreadsheets and saving the result.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Add some initial data and formulas
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=A1+A2"; // B1 will be 30 after calculation
        cells["C1"].Formula = "=B1*2";  // C1 will be 60 after calculation

        // Modify data after setting Manual mode
        cells["A1"].PutValue(15); // Change A1 from 10 to 15
        cells["A2"].PutValue(25); // Change A2 from 20 to 25

        // Ensure full calculation is performed when we trigger it
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;

        // Trigger full recalculation on demand
        workbook.CalculateFormula();

        // Output results to verify calculation
        Console.WriteLine("B1 = " + cells["B1"].Value); // Expected 40
        Console.WriteLine("C1 = " + cells["C1"].Value); // Expected 80

        // Save the workbook (optional)
        workbook.Save("ManualCalculationDemo.xlsx");
    }
}
