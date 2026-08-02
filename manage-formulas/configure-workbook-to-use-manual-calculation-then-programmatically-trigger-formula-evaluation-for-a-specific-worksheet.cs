// Title: Aspose.Cells C# – Set Workbook to Manual Calculation Mode and Run Formula Evaluation
// Description: Shows how to create a workbook, add numeric data, assign a SUM formula, switch the calculation mode to Manual, recalculate formulas programmatically with CalculateFormula, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | Manual calculation mode | CalcModeType.Manual | CalculateFormula | programmatic formula evaluation | worksheet recalculation | Excel automation | performance optimization
// Common Searches: Aspose.Cells set manual calculation mode C# | how to recalculate formulas programmatically Aspose.Cells | CalculateFormula for a single worksheet Aspose.Cells .NET | manual vs automatic calculation Aspose.Cells example | trigger formula evaluation after data update Aspose.Cells
// Developer Intent: Configure a workbook for manual calculation and invoke formula evaluation on demand.
// Use Cases: Populate large workbooks, keep calculation disabled, then compute formulas once after all data is entered to improve speed. | Recalculate formulas only on a specific sheet after modifying its cells, leaving other sheets untouched. | Generate a final report where formulas are evaluated just before saving to guarantee consistent results.
// AI Prompts: Write C# code using Aspose.Cells to set the workbook's calculation mode to Manual and then calculate formulas for a chosen worksheet on demand. | Provide an example that updates cell values, keeps the workbook in Manual mode, calls CalculateFormula only when needed, and saves the file.

using System;
using Aspose.Cells;

namespace AsposeCellsManualCalcDemo
{
    // Shows how to create a workbook, add numeric data, assign a SUM formula, switch the calculation mode to Manual, recalculate formulas programmatically with CalculateFormula, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a formula that sums the three values
                sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

                // Set calculation mode to Manual
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Trigger calculation for the entire workbook (manual mode)
                workbook.CalculateFormula();

                // Display the calculated result
                Console.WriteLine("Result of B1 after manual calculation: " + sheet.Cells["B1"].Value);

                // Save the workbook (lifecycle save rule)
                string outputPath = "ManualCalculationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
