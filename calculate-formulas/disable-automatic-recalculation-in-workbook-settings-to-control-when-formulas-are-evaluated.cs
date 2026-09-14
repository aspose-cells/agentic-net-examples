// Title: Disable automatic formula recalculation and use manual calculation in Aspose.Cells with C#
// AI Prompts: Generate C# code that sets Aspose.Cells workbook CalculationMode to Manual, disables CalculateOnSave, saves the file without evaluating formulas, then later calls CalculateFormula and saves again. | Explain how to configure Aspose.Cells to skip automatic formula evaluation on save and perform explicit calculation later in a .NET application.
// Common Searches: Aspose.Cells C# disable automatic formula calculation before saving workbook | set manual calculation mode in Aspose.Cells .NET and recalculate later | prevent formula recalculation on workbook save using Aspose.Cells | how to use Workbook.CalculateFormula after manual calculation mode in C# | Aspose.Cells turn off CalculateOnSave property example
// Tags: Aspose.Cells manual CalculationMode configuration | disable CalculateOnSave property Aspose.Cells | invoke Workbook.CalculateFormula manually | prevent auto formula recalculation Aspose.Cells | manual formula evaluation workflow C#

using System;
using Aspose.Cells;

namespace DisableAutomaticRecalculationDemo
{
    // The example creates a new workbook, adds sample data and a SUM formula, sets Workbook.Settings.FormulaSettings.CalculationMode to Manual and CalculateOnSave to false to stop automatic recalculation, saves the workbook without evaluating formulas, then explicitly calls Workbook.CalculateFormula and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data and a formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Disable automatic recalculation:
            // 1. Set calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            // 2. Prevent recalculation on save (optional, reinforces manual mode)
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Save the workbook without calculating formulas
            workbook.Save("ManualCalculation.xlsx");

            // At a later point, calculate formulas explicitly if needed
            workbook.CalculateFormula();

            // Save the workbook after manual calculation
            workbook.Save("ManualCalculation_Calculated.xlsx");
        }
    }
}
