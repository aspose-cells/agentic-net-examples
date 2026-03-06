using System;
using Aspose.Cells;

namespace FormulaTracingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Create LoadOptions and disable automatic formula parsing on open
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false; // formulas will be loaded as raw strings

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Display formulas and current values (values are not calculated yet)
            Console.WriteLine("=== Formulas before parsing ===");
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    Console.WriteLine($"Cell {cell.Name}: Formula = \"{cell.Formula}\", Value = {cell.Value}");
                }
            }

            // Parse all formulas that were loaded without parsing
            workbook.ParseFormulas(false); // false => do not ignore errors

            // Calculate all formulas
            workbook.CalculateFormula();

            // Display formulas and calculated values after parsing and calculation
            Console.WriteLine("\n=== Formulas after parsing and calculation ===");
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    Console.WriteLine($"Cell {cell.Name}: Formula = \"{cell.Formula}\", Value = {cell.Value}");
                }
            }

            // Save the workbook to verify that formulas are now stored as calculated values
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to \"{outputPath}\".");
        }
    }
}