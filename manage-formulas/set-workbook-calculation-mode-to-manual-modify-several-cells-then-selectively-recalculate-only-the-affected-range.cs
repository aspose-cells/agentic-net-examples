// Title: Aspose.Cells C# – Set Manual Calculation Mode and Recalculate Specific Cells
// Description: Demonstrates how to switch an Aspose.Cells workbook to manual calculation mode, enable the calculation chain, modify source cells, and selectively recalculate only the dependent formulas (C1 and D1) using CalculateFormula, then save the workbook.
// Keywords: Aspose.Cells manual calculation | C# selective formula recalculation | Enable calculation chain Aspose.Cells | CalculateFormula method example | Performance optimization Excel .NET | Manual mode workbook Aspose | Recalculate specific range Aspose.Cells
// Common Searches: Aspose.Cells set manual calculation mode C# | How to recalculate only changed cells with Aspose.Cells | Enable calculation chain for faster updates Aspose.Cells | CalculateFormula for single formula Aspose.Cells | Manual vs automatic calculation Aspose.Cells
// Developer Intent: The developer wants to improve performance by disabling automatic formula evaluation, update input cells, and then recompute only the formulas that depend on those inputs.
// Use Cases: Speed up large spreadsheets by turning off automatic calculation and updating only affected formulas. | Refresh dependent cells after batch data imports without triggering a full workbook recalculation. | Create interactive reports where user changes affect only a subset of calculations.
// AI Prompts: Generate C# code that sets Aspose.Cells workbook to manual calculation mode, modifies a cell, and uses CalculateFormula to update dependent formulas. | Show how to enable the calculation chain in Aspose.Cells and perform selective recalculation for a range of cells. | Explain the performance benefits of manual calculation mode and how to apply it in a .NET Excel processing workflow.

using System;
using Aspose.Cells;

// Demonstrates how to switch an Aspose.Cells workbook to manual calculation mode, enable the calculation chain, modify source cells, and selectively recalculate only the dependent formulas (C1 and D1) using CalculateFormula, then save the workbook.
class ManualCalculationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set calculation mode to Manual for performance
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Enable calculation chain to help selective recalculation (optional but improves speed)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // -----------------------------------------------------------------
        // Initial data and formulas
        // -----------------------------------------------------------------
        sheet.Cells["A1"].PutValue(10);               // Input value
        sheet.Cells["B1"].PutValue(20);               // Input value
        sheet.Cells["C1"].Formula = "=A1+B1";         // Dependent formula
        sheet.Cells["D1"].Formula = "=C1*2";          // Dependent formula

        // Perform a full calculation once so that all cells have correct values
        workbook.CalculateFormula();

        // -----------------------------------------------------------------
        // Modify some cells
        // -----------------------------------------------------------------
        sheet.Cells["A1"].PutValue(30); // Change only A1

        // -----------------------------------------------------------------
        // Selectively recalculate only the affected range (C1 and D1)
        // -----------------------------------------------------------------
        // Recalculate C1 based on the new A1 value
        object c1Result = sheet.CalculateFormula("=A1+B1");
        sheet.Cells["C1"].PutValue(c1Result);

        // Recalculate D1 based on the updated C1 value
        object d1Result = sheet.CalculateFormula("=C1*2");
        sheet.Cells["D1"].PutValue(d1Result);

        // -----------------------------------------------------------------
        // Save the workbook
        // -----------------------------------------------------------------
        workbook.Save("ManualCalculationDemo.xlsx");
    }
}
