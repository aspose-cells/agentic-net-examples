// Title: Disable Auto-Calc and Manually Recalculate Formulas per Worksheet in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set Aspose.Cells workbook calculation mode to Manual, insert static values and formulas on multiple sheets, and use Worksheet.CalculateFormula with CalculationOptions to evaluate formulas only on selected worksheets before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | CalcModeType.Manual | Worksheet.CalculateFormula | calculate formulas per sheet | disable automatic formula evaluation | Excel performance optimization | multi‑sheet workbook | custom CalculationOptions
// Common Searches: Aspose.Cells disable automatic calculation | Worksheet.CalculateFormula C# example | manual formula evaluation per worksheet Aspose | set CalcModeType.Manual in .NET | how to recalculate specific sheets with Aspose.Cells
// Developer Intent: Turn off auto‑calc, add formulas, and trigger calculation only on chosen worksheets.
// Use Cases: Large financial models where formulas are evaluated after all input data is entered. | Generating multi‑sheet reports that require independent recalculation after each sheet is populated. | Improving workbook creation speed by disabling auto‑calc and manually invoking calculation on demand.
// AI Prompts: Show C# code that sets CalcModeType.Manual, adds formulas to several worksheets, and uses Worksheet.CalculateFormula with custom CalculationOptions to recalculate each sheet individually. | Provide an Aspose.Cells example that disables automatic formula calculation, inserts formulas, and manually triggers calculation for specific worksheets before saving the file.

using System;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells workbook calculation mode to Manual, insert static values and formulas on multiple sheets, and use Worksheet.CalculateFormula with CalculationOptions to evaluate formulas only on selected worksheets before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (uses the provided create rule)
        Workbook workbook = new Workbook();

        // Disable automatic calculation for the whole workbook
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // -------------------------------------------------
        // Worksheet 1 – add data and a formula
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Put a static value
        sheet1.Cells["A1"].PutValue(5);
        // Insert a formula that depends on A1
        sheet1.Cells["B1"].Formula = "=A1*2";

        // -------------------------------------------------
        // Worksheet 2 – add data and a different formula
        // -------------------------------------------------
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        sheet2.Cells["A1"].PutValue(10);
        sheet2.Cells["B1"].Formula = "=A1+100";

        // -------------------------------------------------
        // Manually trigger calculation for each worksheet
        // -------------------------------------------------
        CalculationOptions calcOptions = new CalculationOptions();

        // Calculate all formulas in Sheet1
        sheet1.CalculateFormula(calcOptions, true);

        // Calculate all formulas in Sheet2
        sheet2.CalculateFormula(calcOptions, true);

        // Save the workbook (uses the provided save rule)
        workbook.Save("ManualCalculationDemo.xlsx");
    }
}
