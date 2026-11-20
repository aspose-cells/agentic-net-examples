// Title: Calculate all formulas in an Excel workbook using Aspose.Cells Workbook.CalculateFormula with default settings in C#
// Common Searches: C# Aspose.Cells calculate all worksheet formulas with default options | How to evaluate Excel formulas using Workbook.CalculateFormula in .NET | Run default formula calculation on a loaded workbook using Aspose.Cells | Compute and save Excel workbook after formula evaluation with Aspose.Cells C# | Use Aspose.Cells to recalculate formulas without custom calculation settings
// Tags: Workbook.CalculateFormula default mode | Aspose.Cells recalculate all workbook formulas | C# load Excel workbook and compute formulas | persist workbook after formula calculation | default calculation mode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads an existing Excel file into an Aspose.Cells Workbook (optionally using LoadOptions for password‑protected files), invokes Workbook.CalculateFormula() to evaluate every formula with the library's default calculation engine, and saves the updated workbook, handling both Aspose.Cells‑specific and generic exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook. If the file is password‑protected, set the password in LoadOptions.
                var loadOptions = new LoadOptions();
                // loadOptions.Password = "yourPassword"; // Uncomment and set if needed

                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Calculate all formulas using default settings
                workbook.CalculateFormula();

                // Save the workbook after calculation
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handles Aspose.Cells specific errors (e.g., missing password)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handles any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
