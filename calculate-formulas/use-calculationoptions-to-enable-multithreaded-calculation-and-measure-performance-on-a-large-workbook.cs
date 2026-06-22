using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

class MultiThreadedCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Define size of the large workbook
            int rows = 5000;   // number of rows
            int cols = 100;    // number of data columns (A..CV)

            // Populate cells with numeric data and add a SUM formula at the end of each row
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j].PutValue(i + j);
                }

                // Formula column (after the data columns) sums the whole row
                string lastColName = GetColumnName(cols);
                cells[i, cols].Formula = $"=SUM(A{i + 1}:{lastColName}{i + 1})";
            }

            // Enable multi‑threaded calculation if the API supports it
            // (property may not exist in older versions; safe to omit)
            // wb.Settings.EnableMultiThreadedCalculation = true;

            // Set calculation options (ignore errors and calculate recursively)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true,
                Recursive = true
            };

            // Measure calculation performance
            Stopwatch sw = Stopwatch.StartNew();
            wb.CalculateFormula(calcOptions);
            sw.Stop();

            Console.WriteLine($"Multi‑threaded calculation completed in {sw.ElapsedMilliseconds} ms");

            // Save the workbook (ensure the directory exists)
            string outputPath = "LargeWorkbook.xlsx";
            try
            {
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to convert a 1‑based column index to an Excel column name (e.g., 1 -> A, 27 -> AA)
    static string GetColumnName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = "";
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            name = letters[rem] + name;
            index = (index - 1) / 26;
        }
        return name;
    }
}