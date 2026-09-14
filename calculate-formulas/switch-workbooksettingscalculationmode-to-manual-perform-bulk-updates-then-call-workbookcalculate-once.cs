// Title: Set Aspose.Cells workbook to Manual calculation mode, bulk‑populate cells, then run a single CalculateFormula in C#
// AI Prompts: Demonstrate how to turn off automatic formula recalculation, efficiently fill a large cell range, add dependent formulas, and invoke Workbook.CalculateFormula once using Aspose.Cells for .NET. | Provide C# code that sets Workbook.Settings.FormulaSettings.CalculationMode to Manual, performs high‑volume cell updates, and triggers a single formula evaluation.
// Common Searches: asp.net aspose.cells manual calculation mode for bulk data import | c# fill thousands of cells in worksheet then calculate formulas once | how to improve performance when inserting large ranges with Aspose.Cells | disable automatic recalculation Aspose.Cells before bulk update | example of using CalcModeType.Manual with SUM formulas in Aspose.Cells
// Tags: manual formula evaluation Aspose.Cells | bulk cell insertion .NET | CalculateFormula after large data load | FormulaSettings.CalcModeType.Manual usage | high‑volume worksheet population Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, switches calculation to Manual, efficiently fills a 1000 × 10 range, adds SUM formulas, runs a single CalculateFormula call, and saves the file as BulkUpdateManualCalc.xlsx.
class BulkUpdateWithManualCalculation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Perform bulk updates (example: fill a 1000x10 range with values)
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                cells[row, col].PutValue(row + col);
            }
        }

        // Add some formulas that depend on the bulk data
        cells["K1"].Formula = "=SUM(A1:J1)";
        cells["K2"].Formula = "=SUM(A2:J2)";
        cells["K3"].Formula = "=SUM(A3:J3)";

        // Calculate all formulas once
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("BulkUpdateManualCalc.xlsx");
    }
}
