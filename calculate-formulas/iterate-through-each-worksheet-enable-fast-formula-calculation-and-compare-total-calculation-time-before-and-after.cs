using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Ensure the workbook has at least 5 worksheets
            const int requiredSheets = 5;
            while (workbook.Worksheets.Count < requiredSheets)
            {
                workbook.Worksheets.Add();
            }

            // Populate worksheets with a large number of formulas
            for (int sheetIndex = 0; sheetIndex < requiredSheets; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                for (int row = 0; row < 1000; row++)
                {
                    // Simple numeric value in column A
                    sheet.Cells[row, 0].PutValue(row);
                    // Formula that depends on the value in column A
                    sheet.Cells[row, 1].Formula = $"=A{row + 1}+10";
                }
            }

            // -------------------------------------------------
            // First calculation: fast formula calculation disabled
            // -------------------------------------------------
            workbook.Settings.FormulaSettings.EnableCalculationChain = false;

            Stopwatch timer = new Stopwatch();
            timer.Start();

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            timer.Stop();
            Console.WriteLine($"Calculation time without chain: {timer.ElapsedMilliseconds} ms");

            // -------------------------------------------------
            // Enable fast formula calculation (calculation chain)
            // -------------------------------------------------
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Re‑calculate and measure time again
            timer.Restart();

            workbook.CalculateFormula();

            timer.Stop();
            Console.WriteLine($"Calculation time with chain: {timer.ElapsedMilliseconds} ms");

            // Save the workbook
            string outputPath = "FastFormulaDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}