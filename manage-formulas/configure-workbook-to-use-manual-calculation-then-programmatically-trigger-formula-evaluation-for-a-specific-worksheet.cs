// Title: Aspose.Cells C# – Set Manual Calculation Mode and Recalculate a Single Worksheet
// Description: Shows how to switch a workbook to manual calculation, configure optional CalculationOptions, and call worksheet.CalculateFormula to evaluate formulas only on the chosen sheet before saving the file.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | CalcModeType.Manual | CalculateFormula | recalculate single worksheet | disable recursive calculation | formula evaluation | Aspose.Cells example | GitHub Aspose.Cells manual calc
// Common Searches: Aspose.Cells set manual calculation mode C# | CalculateFormula for one worksheet only | disable recursive formula calculation Aspose.Cells | trigger formula evaluation after cell update Aspose.Cells .NET | manual calc mode example Aspose.Cells GitHub
// Developer Intent: The developer needs to change the workbook’s calculation setting to manual and programmatically recalculate formulas on a specific worksheet without affecting other sheets.
// Use Cases: Recompute formulas on a large sheet after bulk data changes while keeping other sheets untouched. | Generate on‑demand reports where formula evaluation is deferred for performance reasons. | Perform iterative or conditional calculations on a single worksheet without triggering full‑workbook recalculation.
// AI Prompts: Provide C# code that sets CalcModeType.Manual, updates cells, and calls worksheet.CalculateFormula with CalculationOptions to recalculate only that sheet. | Explain how to disable recursive calculation when using Aspose.Cells to evaluate formulas on a specific worksheet. | Show an example of triggering manual formula evaluation for a worksheet, retrieving the result, and saving the workbook.

using System;
using Aspose.Cells;

// Shows how to switch a workbook to manual calculation, configure optional CalculationOptions, and call worksheet.CalculateFormula to evaluate formulas only on the chosen sheet before saving the file.
class ManualCalculationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        // Add a formula that depends on the above cells
        worksheet.Cells["B1"].Formula = "=A1+A2";

        // Set the workbook's calculation mode to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Prepare calculation options (optional settings)
        CalculationOptions calcOptions = new CalculationOptions
        {
            // Enable recursive calculation only if you want dependent worksheets to be processed
            Recursive = false
        };

        // Trigger calculation for the specific worksheet
        // The second parameter (recursive) is set to false to limit calculation to this worksheet only
        worksheet.CalculateFormula(calcOptions, false);

        // Output the calculated result to verify
        Console.WriteLine("B1 value after manual calculation: " + worksheet.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("ManualCalculationResult.xlsx");
    }
}
