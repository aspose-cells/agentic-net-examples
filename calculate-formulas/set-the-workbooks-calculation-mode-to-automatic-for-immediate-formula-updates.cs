// Title: Set Aspose.Cells workbook to Automatic calculation mode in C# for instant formula updates
// AI Prompts: Generate C# code that configures a Workbook's FormulaSettings to use Automatic calculation mode and saves the file with Aspose.Cells. | Demonstrate adding values, assigning a formula, invoking calculation, and persisting the workbook after enabling automatic recalculation using Aspose.Cells in .NET.
// Common Searches: Aspose.Cells how to set calculation mode to automatic in C# | C# example of automatic formula recalculation using Aspose.Cells | Enable immediate formula updates in Aspose.Cells workbook .NET | Set CalcModeType.Automatic for a workbook with Aspose.Cells | Automatic calculation of formulas after cell changes Aspose.Cells C#
// Tags: automatic calculation mode Aspose.Cells | set CalcModeType.Automatic C# | Aspose.Cells workbook formula auto-recalculation | C# save workbook after automatic calculation | FormulaSettings.CalculationMode usage Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, switches its calculation mode to Automatic for real‑time formula evaluation, adds sample data and a simple addition formula, forces calculation, and saves the result as AutomaticCalculation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the calculation mode to Automatic for immediate formula updates
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Add sample data and a formula to demonstrate automatic calculation
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].Formula = "=A1+A2";

        // Since the mode is Automatic, we can calculate now to get the result
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("AutomaticCalculation.xlsx");
    }
}
