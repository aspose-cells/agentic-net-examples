using Aspose.Cells;
using System;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable iterative calculation to resolve circular references
                var formulaSettings = workbook.Settings.FormulaSettings;
                formulaSettings.EnableIterativeCalculation = true;
                formulaSettings.MaxIteration = 100;
                formulaSettings.MaxChange = 0.001; // convergence threshold

                // Example circular reference
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1+1";
                sheet.Cells["B1"].Formula = "=A1+1";

                // Calculate all formulas (iterative settings are applied automatically)
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "IterativeCalculation.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}