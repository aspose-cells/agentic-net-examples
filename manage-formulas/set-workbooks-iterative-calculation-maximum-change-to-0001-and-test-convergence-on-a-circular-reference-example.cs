using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class IterativeCalculationMaxChangeDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Enable iterative calculation and set thresholds
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;   // optional, ensures enough iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;   // maximum change to stop iteration

                // Create a circular reference:
                // A1 = B1 + 1
                // B1 = A1 + 1
                cells["A1"].Formula = "=B1+1";
                cells["B1"].Formula = "=A1+1";

                // Perform calculation with the iterative settings
                workbook.CalculateFormula();

                // Output the results – they should converge within the MaxChange limit
                Console.WriteLine("A1 value after calculation: " + cells["A1"].Value);
                Console.WriteLine("B1 value after calculation: " + cells["B1"].Value);
                Console.WriteLine("Maximum change used: " + workbook.Settings.FormulaSettings.MaxChange);

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "IterativeCalculationMaxChangeDemo.xlsx");

                // Save the workbook (optional, demonstrates that settings are persisted)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions (e.g., calculation errors, I/O issues)
                Console.WriteLine($"Error during execution: {ex.Message}");
            }
        }
    }
}