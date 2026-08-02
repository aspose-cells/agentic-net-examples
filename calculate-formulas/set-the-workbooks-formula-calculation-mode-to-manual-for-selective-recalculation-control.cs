// Title: Aspose.Cells C# – Set Workbook Formula Calculation Mode to Manual
// Description: Demonstrates how to switch a workbook to manual formula calculation in Aspose.Cells for .NET, prevent automatic recalculation, trigger a single manual evaluation with CalculateFormula, and save the result.
// Keywords: Aspose.Cells manual calculation | C# set calculation mode | Workbook.Settings.FormulaSettings | CalcModeType.Manual | disable automatic formula evaluation | manual formula recalculation .NET | performance optimization Aspose.Cells
// Common Searches: Aspose.Cells set manual calculation mode | how to disable automatic formula recalculation C# | manual formula evaluation Aspose.Cells .NET | calculate formulas on demand Aspose.Cells
// Developer Intent: Switch the workbook’s formula engine to manual mode and run calculations only when explicitly requested.
// Use Cases: Populate a large financial model, keep formulas unevaluated during data entry, then compute once with CalculateFormula. | Create a template that users will fill later; avoid premature calculations to improve generation speed. | Process many worksheets in a batch, suppress repeated recalculations, and invoke a single manual calculation after all updates.
// AI Prompts: Show C# code to toggle between Manual and Automatic calculation modes in Aspose.Cells and recalculate a specific worksheet. | Explain how Manual calculation mode improves performance when loading massive datasets with formulas in Aspose.Cells. | Provide an example of using workbook.CalculateFormula after bulk cell updates while the workbook is in manual mode.

using System;
using Aspose.Cells;

// Demonstrates how to switch a workbook to manual formula calculation in Aspose.Cells for .NET, prevent automatic recalculation, trigger a single manual evaluation with CalculateFormula, and save the result.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data and a formula
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(10);
        cells["B1"].Formula = "=A1+A2";

        // Set the calculation mode to Manual for selective recalculation
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // At this point the formula result is not calculated automatically
        Console.WriteLine("Before manual calculation, B1 value: " + cells["B1"].Value);

        // Perform manual calculation when required
        workbook.CalculateFormula();

        Console.WriteLine("After manual calculation, B1 value: " + cells["B1"].IntValue);

        // Save the workbook
        workbook.Save("ManualCalculationMode.xlsx");
    }
}
