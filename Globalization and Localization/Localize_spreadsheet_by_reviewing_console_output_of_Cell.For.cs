using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale to German to see localized formulas (e.g., SUM -> SUMME)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate over all used cells
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Standard (English) formula
                        string standardFormula = cell.Formula;

                        // Locale‑formatted formula (German in this case)
                        string localizedFormula = cell.FormulaLocal;

                        // GetFormula with explicit locale flag (should match FormulaLocal)
                        string getFormulaLocal = cell.GetFormula(false, true);

                        // Output the information to the console
                        Console.WriteLine($"Cell {cell.Name}:");
                        Console.WriteLine($"  Standard Formula : {standardFormula}");
                        Console.WriteLine($"  Localized Formula: {localizedFormula}");
                        Console.WriteLine($"  GetFormula(true) : {getFormulaLocal}");
                        Console.WriteLine();
                    }
                }
            }

            // Optionally save the workbook (unchanged) to demonstrate lifecycle compliance
            workbook.Save("output.xlsx");
        }
    }
}