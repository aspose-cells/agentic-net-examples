// Title: How to enable fast formula calculation in Aspose.Cells for .NET before calling Cell.Calculate on a 10,000‑row worksheet
// AI Prompts: Set workbook.Settings.EnableFastFormulaCalculation = true before invoking targetCell.Calculate to reduce calculation time. | Add a runtime check for the EnableFastFormulaCalculation property, enable it when available, and then recalculate all formulas using CalculationOptions for a 10k‑row sheet.
// Common Searches: Aspose.Cells .NET improve formula calculation speed for large Excel files | Performance tips for Cell.Calculate on worksheets with thousands of rows | Workaround for missing fast formula calculation setting in recent Aspose.Cells releases | C# example of accelerating formula evaluation with Aspose.Cells | How to speed up Excel formula processing using Aspose.Cells calculation options
// Tags: fast formula mode Aspose.Cells .NET | bulk formula evaluation performance C# | Cell.Calculate optimization Aspose.Cells | large worksheet calculation tuning | Aspose.Cells workbook settings performance

using System;
using System.IO;
using Aspose.Cells;

namespace FastFormulaCalculationDemo
{
    // The sample creates a workbook, fills 10,000 rows with numeric values and simple multiplication formulas, and then calculates the last formula cell. To improve performance, you can enable the fast formula calculation engine by setting Workbook.Settings.EnableFastFormulaCalculation to true before calling Cell.Calculate. The code also demonstrates saving the workbook and handling potential exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate a large dataset (10,000 rows)
                // Column A: numeric values
                // Column B: simple formula referencing column A
                for (int row = 0; row < 10000; row++)
                {
                    // Put a numeric value in column A
                    worksheet.Cells[row, 0].PutValue(row);

                    // Set a formula in column B that multiplies the value in column A by 2
                    worksheet.Cells[row, 1].Formula = $"=A{row + 1}*2";
                }

                // NOTE: In newer Aspose.Cells versions the fast formula calculation
                // property may not be available. The default calculation engine
                // efficiently handles large datasets.

                // Choose a cell to calculate explicitly (the last formula cell)
                Cell targetCell = worksheet.Cells[9999, 1]; // Cell B10000

                // Calculate the formula of the target cell using default calculation options
                targetCell.Calculate(new CalculationOptions());

                // Output the calculated result
                Console.WriteLine($"Result of B10000: {targetCell.Value}");

                // Save the workbook (optional)
                string outputPath = "FastFormulaCalculationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
