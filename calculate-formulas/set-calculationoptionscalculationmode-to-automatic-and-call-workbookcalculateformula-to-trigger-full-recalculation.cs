// Title: Enable Automatic Calculation Mode and Force Full Formula Recalculation with Aspose.Cells in C#
// AI Prompts: Set workbook.FormulaSettings.CalculationMode to CalcModeType.Automatic and invoke Workbook.CalculateFormula with a CalculationOptions object to recalculate every formula. | Programmatically trigger a complete workbook recalculation after changing the calculation mode using the Aspose.Cells .NET API. | Use CalculationOptions to perform a full formula evaluation across all worksheets in a newly created workbook.
// Common Searches: asp.net aspose.cells set calculation mode to automatic and recalculate formulas | c# how to force full workbook calculation after changing CalcModeType in Aspose.Cells | using CalculationOptions with Workbook.CalculateFormula to update all cells Aspose.Cells | trigger automatic formula evaluation in Aspose.Cells .NET example
// Tags: Aspose.Cells automatic calculation mode | Workbook.CalculateFormula full recalculation | CalculationOptions Aspose.Cells usage | set CalcModeType Automatic C# | recalculate all formulas Aspose.Cells | FormulaSettings.CalculationMode .NET

using System;
using Aspose.Cells;

namespace AsposeCellsCalculationDemo
{
    // The example creates a new workbook, adds numeric values and a SUM formula, switches the calculation mode to Automatic via FormulaSettings, creates a default CalculationOptions instance, calls Workbook.CalculateFormula to recalculate all formulas, prints the result of the SUM cell, and saves the workbook as CalculationResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["B1"].Formula = "=SUM(A1:A3)"; // B1 will hold the sum of A1:A3

            // Set the calculation mode to Automatic (saved in the file for Excel)
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // Create calculation options (default options are sufficient)
            CalculationOptions calcOptions = new CalculationOptions();

            // Trigger full recalculation of all formulas in the workbook
            workbook.CalculateFormula(calcOptions);

            // Verify that the formula result is calculated
            Console.WriteLine("Result of B1 (SUM): " + cells["B1"].Value);

            // Save the workbook to a file
            workbook.Save("CalculationResult.xlsx");
        }
    }
}
