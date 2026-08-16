// Title: C# – Set Automatic Calculation Mode and Recalculate All Formulas with Aspose.Cells
// Description: Shows how to enable automatic formula calculation (CalcModeType.Automatic) in Aspose.Cells for .NET, create a CalculationOptions instance, call Workbook.CalculateFormula for a full workbook recalculation, read the computed values, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | automatic calculation mode | CalcModeType.Automatic | Workbook.CalculateFormula | CalculationOptions | recalculate formulas | full workbook recalculation | Excel automation | save workbook
// Common Searches: Aspose.Cells set calculation mode to automatic C# | How to force full recalculation of formulas in Aspose.Cells | Workbook.CalculateFormula example .NET | Using CalculationOptions with Aspose.Cells | Enable automatic formula evaluation Aspose.Cells
// Developer Intent: Enable automatic formula evaluation and trigger a complete workbook recalculation.
// Use Cases: Refresh all dependent cells after programmatically changing input values before exporting the file. | Generate Excel reports with formulas and guarantee that every result is up‑to‑date by invoking a full recalculation. | Validate complex spreadsheet logic in automated tests by forcing Aspose.Cells to recompute every formula.
// AI Prompts: Provide a C# snippet that sets CalcModeType.Automatic and runs Workbook.CalculateFormula in Aspose.Cells. | How can I force a full workbook recalculation after updating cells using Aspose.Cells for .NET? | Explain the steps to retrieve calculated values after calling Workbook.CalculateFormula with CalculationOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationDemo
{
    // Shows how to enable automatic formula calculation (CalcModeType.Automatic) in Aspose.Cells for .NET, create a CalculationOptions instance, call Workbook.CalculateFormula for a full workbook recalculation, read the computed values, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=A1+A2";          // Simple addition
            cells["B1"].Formula = "=SUM(A1:A2)";    // Sum function

            // Set the calculation mode to Automatic (via FormulaSettings)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Create calculation options (default options)
            CalculationOptions calcOptions = new CalculationOptions();

            // Trigger full recalculation of all formulas in the workbook
            workbook.CalculateFormula(calcOptions);

            // Display the calculated results
            Console.WriteLine("A3 (A1+A2) = " + cells["A3"].Value);
            Console.WriteLine("B1 (SUM(A1:A2)) = " + cells["B1"].Value);

            // Save the workbook (using the standard save rule)
            workbook.Save("CalculatedWorkbook.xlsx");
        }
    }
}
