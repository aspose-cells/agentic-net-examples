// Title: Manual Calculation Mode with Forced Full Recalculation in Aspose.Cells for .NET (C#)
// Description: Shows how to set a workbook’s FormulaSettings.CalculationMode to Manual, modify cell values, enable ForceFullCalculation, invoke CalculateFormula, and save the spreadsheet using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | CalcModeType.Manual | ForceFullCalculation | CalculateFormula | Excel formula recalculation | batch cell updates | performance optimization
// Common Searches: Aspose.Cells set calculation mode manual | force full calculation Aspose.Cells C# | CalculateFormula after cell updates .NET | disable automatic formula evaluation Aspose.Cells | manual workbook recalculation example
// Developer Intent: Configure a workbook to use manual calculation, change data without triggering automatic updates, then run a single full formula evaluation on demand.
// Use Cases: Boost performance when generating large spreadsheets by postponing formula evaluation until all data is inserted. | Apply bulk data transformations and guarantee consistent results with one forced recalculation before exporting the file. | Produce deterministic Excel reports where formulas are evaluated only after all input cells have been set.
// AI Prompts: Provide C# code that sets CalcModeType.Manual, updates cells A1 and A2, enables ForceFullCalculation, calls CalculateFormula, and saves the workbook as ManualCalcDemo.xlsx using Aspose.Cells. | Show an example that disables automatic calculation, performs batch updates on a worksheet, then triggers a full recalculation in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to set a workbook’s FormulaSettings.CalculationMode to Manual, modify cell values, enable ForceFullCalculation, invoke CalculateFormula, and save the spreadsheet using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set calculation mode to Manual (no automatic recalculation)
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Add initial data and a formula that depends on the data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["B1"].Formula = "=A1+A2"; // simple sum formula

        // Modify the data after setting manual mode
        sheet.Cells["A1"].PutValue(30);
        sheet.Cells["A2"].PutValue(40);

        // Ensure a full calculation is performed each time we trigger it
        workbook.Settings.FormulaSettings.ForceFullCalculation = true;

        // Trigger full recalculation on demand
        workbook.CalculateFormula();

        // Verify the result of the formula after recalculation
        Console.WriteLine("B1 value after recalculation: " + sheet.Cells["B1"].Value);

        // Save the workbook (lifecycle: save)
        workbook.Save("ManualCalcDemo.xlsx");
    }
}
