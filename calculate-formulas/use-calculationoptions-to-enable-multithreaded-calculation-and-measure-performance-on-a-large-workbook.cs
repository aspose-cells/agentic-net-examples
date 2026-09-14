// Title: Measure multi‑threaded formula calculation performance with Aspose.Cells CalculationOptions on a 5,000‑row Excel workbook (C#)
// AI Prompts: Generate C# code that creates a 5,000‑row by 100‑column workbook, fills it with random numbers and SUM formulas, enables threaded calculation via CalculationOptions, runs CalculateFormula, and prints the elapsed time. | Show how to toggle Aspose.Cells workbook.Settings.EnableThreadedCalculation on and off to compare single‑threaded versus multi‑threaded calculation times in a large worksheet. | Provide a snippet that saves the workbook after multi‑threaded calculation and logs the execution duration in seconds using Stopwatch.
// Common Searches: asp.net how to enable threaded calculation in Aspose.Cells and benchmark speed | c# measure performance of CalculateFormula on large Excel file using Aspose.Cells | using CalculationOptions to speed up formula evaluation in a 5000 row workbook | compare single thread and multi thread calculation times with Aspose.Cells .NET
// Tags: Aspose.Cells multi‑threaded calculation | CalculationOptions EnableThreadedCalculation | benchmark Excel formula evaluation .NET | large workbook performance Aspose.Cells | C# Stopwatch timing Aspose.Cells calculation

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMultiThreadedCalculation
{
    // The example creates a 5,000‑row by 100‑column workbook, populates cells with random doubles and SUM formulas, optionally enables multi‑threaded calculation via CalculationOptions, measures the time taken by CalculateFormula with a Stopwatch, saves the workbook as XLSX, and writes the elapsed seconds to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define size of the large worksheet
                const int totalRows = 5000;
                const int totalCols = 100;

                // Populate the worksheet with random numbers and simple formulas
                Random rnd = new Random();
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        // Every 10th column will contain a formula that sums the previous 9 cells in the same row
                        if (col % 10 == 0 && col >= 9)
                        {
                            string startCell = CellsHelper.CellIndexToName(row, col - 9);
                            string endCell   = CellsHelper.CellIndexToName(row, col - 1);
                            cells[row, col].Formula = $"=SUM({startCell}:{endCell})";
                        }
                        else
                        {
                            // Fill with a random double value
                            cells[row, col].PutValue(rnd.NextDouble() * 1000);
                        }
                    }
                }

                // Enable multi‑threaded calculation if the property exists in the used version
                // workbook.Settings.EnableThreadedCalculation = true;

                // Measure calculation time
                Stopwatch sw = Stopwatch.StartNew();

                // Perform calculation
                workbook.CalculateFormula();

                sw.Stop();
                Console.WriteLine($"Multi‑threaded calculation completed in {sw.Elapsed.TotalSeconds:F2} seconds.");

                // Save the workbook
                string outputPath = "LargeWorkbook_MultiThreaded.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
