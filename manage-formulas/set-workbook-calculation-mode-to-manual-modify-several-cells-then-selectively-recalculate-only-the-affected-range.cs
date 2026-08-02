// Title: Aspose.Cells .NET: Manual Calculation Mode with Selective Recalculation (C# Example)
// Description: Demonstrates how to set a workbook to manual calculation, enable the calculation chain, modify input cells, and trigger CalculateFormula so that only dependent formulas are recomputed before saving the file.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | selective formula recalculation | EnableCalculationChain | CalcModeType.Manual | performance optimization | partial workbook calculation | Excel automation
// Common Searches: Aspose.Cells set manual calculation mode C# | How to recalculate only dependent cells with Aspose.Cells | EnableCalculationChain example .NET | Selective formula evaluation Aspose.Cells | CalculateFormula specific range C# | Performance tips for large spreadsheets Aspose.Cells
// Developer Intent: Switch workbook to manual calculation, update cells, and recalculate only affected formulas.
// Use Cases: Speed up batch updates in large financial models by recalculating only changed formulas. | Create reporting tools that modify input data and need quick partial recalculation. | Implement server‑side Excel processing where full workbook calculation would be costly. | Run automated data imports that adjust a subset of cells without triggering full recompute.
// AI Prompts: Provide a C# snippet that sets Aspose.Cells workbook to manual mode, enables the calculation chain, updates cells, and triggers selective recalculation. | Explain the impact of EnableCalculationChain on performance when using CalculateFormula in Aspose.Cells. | Show how to recalculate a defined cell range after changes while keeping the rest of the workbook untouched. | Give guidance on switching back to automatic calculation after selective updates in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set a workbook to manual calculation, enable the calculation chain, modify input cells, and trigger CalculateFormula so that only dependent formulas are recomputed before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -----------------------------------------------------------------
        // Set up initial data and formulas
        // -----------------------------------------------------------------
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);

        // Formulas that depend on the values in column A
        sheet.Cells["B1"].Formula = "=A1*10";
        sheet.Cells["B2"].Formula = "=A2*10";
        sheet.Cells["C1"].Formula = "=SUM(B1:B2)";

        // -----------------------------------------------------------------
        // Configure calculation settings
        // -----------------------------------------------------------------
        // Switch to manual calculation mode
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Enable calculation chain so that only dependent cells are recalculated
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Perform an initial calculation to build the chain
        workbook.CalculateFormula();

        // -----------------------------------------------------------------
        // Modify cells that affect formulas
        // -----------------------------------------------------------------
        sheet.Cells["A1"].PutValue(5); // Affects B1 and consequently C1
        sheet.Cells["A2"].PutValue(7); // Affects B2 and consequently C1

        // Recalculate only the cells that depend on the changed values
        workbook.CalculateFormula();

        // -----------------------------------------------------------------
        // Save the workbook
        // -----------------------------------------------------------------
        workbook.Save("ManualCalcSelective.xlsx");
    }
}
