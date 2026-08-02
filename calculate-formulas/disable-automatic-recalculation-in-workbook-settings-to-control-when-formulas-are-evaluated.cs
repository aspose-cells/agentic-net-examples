// Title: Disable Automatic Formula Recalculation in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set Aspose.Cells Workbook.Settings.FormulaSettings to Manual mode and turn off CalculateOnSave, so formulas remain unevaluated until you explicitly recalculate them.
// Keywords: Aspose.Cells manual calculation mode | disable automatic recalculation C# | FormulaSettings CalculationMode Manual | CalculateOnSave false Aspose.Cells | prevent formula evaluation .NET | Aspose.Cells performance optimization
// Common Searches: Aspose.Cells turn off automatic calculation | C# set workbook calculation mode to manual | disable calculate on save Aspose.Cells | how to keep formulas unevaluated in Excel using Aspose | manual formula recalculation Aspose.Cells example
// Developer Intent: The developer wants to stop formulas from being calculated automatically so they can control when evaluation occurs.
// Use Cases: Create a template workbook where formulas stay intact until a later processing step. | Improve generation speed for large spreadsheets by deferring formula evaluation. | Distribute Excel files to end‑users with formulas preserved but not pre‑calculated.
// AI Prompts: Show how to enable manual calculation mode, add formulas, and later trigger a full workbook recalc in Aspose.Cells for .NET. | Provide C# code to toggle automatic recalculation on or off based on a runtime flag using Aspose.Cells settings. | Explain how to recalculate a specific worksheet after setting CalculationMode to Manual in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells Workbook.Settings.FormulaSettings to Manual mode and turn off CalculateOnSave, so formulas remain unevaluated until you explicitly recalculate them.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Disable automatic recalculation by setting the mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prevent formulas from being recalculated when the workbook is saved
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Add sample data and a formula (optional, just for demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Save the workbook; formulas remain unevaluated
        workbook.Save("output_without_calculation.xlsx");
    }
}
