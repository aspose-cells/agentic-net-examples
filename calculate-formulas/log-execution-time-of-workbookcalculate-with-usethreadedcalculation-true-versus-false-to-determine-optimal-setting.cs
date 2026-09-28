// Title: Benchmark Aspose.Cells Workbook.CalculateFormula with threaded calculation enabled and disabled in C#
// AI Prompts: Create a C# program that populates a large worksheet with random numbers, adds row‑wise SUM formulas, and records the elapsed time of Workbook.CalculateFormula when EnableThreadedCalculation is true. | Demonstrate using reflection to set Workbook.Settings.EnableThreadedCalculation to false, run the calculation again, and log both execution times for side‑by‑side comparison.
// Common Searches: Aspose.Cells calculateformula performance with EnableThreadedCalculation true vs false | C# benchmark multi‑threaded formula calculation in Aspose.Cells | how to turn off threaded calculation in Aspose.Cells for large Excel files | measure execution time of Aspose.Cells workbook calculation in .NET | optimal threading setting for Aspose.Cells formula evaluation on 10,000 rows
// Tags: Aspose.Cells workbook calculation performance | EnableThreadedCalculation reflection toggle | multi‑threaded formula evaluation Aspose.Cells | large worksheet benchmark C# | Aspose.Cells calculation speed comparison

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example builds a 10,000‑row worksheet with random numeric data, adds a SUM formula to each row, performs an initial calculation, then measures the time taken by Workbook.CalculateFormula with the default (threaded) setting and again after disabling threading via reflection, outputting both durations for performance analysis.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and populate it with sample data and formulas
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                int rows = 10000;
                int cols = 10;
                Random rnd = new Random();

                // Fill cells with random numbers
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        cells[i, j].PutValue(rnd.NextDouble() * 100);
                    }
                }

                // Add a column with a formula that sums each row
                for (int i = 0; i < rows; i++)
                {
                    cells[i, cols].Formula = $"=SUM(A{i + 1}:{GetColumnLetter(cols)}{i + 1})";
                }

                // Initial calculation to ensure formulas are ready
                workbook.CalculateFormula();

                // Measure calculation time (default settings)
                Stopwatch swDefault = Stopwatch.StartNew();
                workbook.CalculateFormula();
                swDefault.Stop();

                // Attempt to disable threaded calculation via reflection (property may not exist in older versions)
                try
                {
                    var prop = workbook.Settings.GetType().GetProperty("EnableThreadedCalculation");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(workbook.Settings, false);
                    }
                }
                catch
                {
                    // Ignore any reflection errors and continue with default settings
                }

                // Measure calculation time after attempting to disable threading
                Stopwatch swSingleThread = Stopwatch.StartNew();
                workbook.CalculateFormula();
                swSingleThread.Stop();

                // Output the results
                Console.WriteLine($"Calculation time with default settings: {swDefault.ElapsedMilliseconds} ms");
                Console.WriteLine($"Calculation time after attempting to disable threading: {swSingleThread.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert column index to Excel column letter (0‑based index)
        static string GetColumnLetter(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string result = "";
            int dividend = index + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                result = letters[modulo] + result;
                dividend = (dividend - modulo) / 26;
            }
            return result;
        }
    }
}
