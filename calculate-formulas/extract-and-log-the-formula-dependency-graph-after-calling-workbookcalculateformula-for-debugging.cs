// Title: Extracting and logging the Excel formula dependency graph after Workbook.CalculateFormula with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells get formula dependency graph after CalculateFormula | how to list cells that depend on a formula using Aspose.Cells .NET | debug Excel formulas by printing dependency information with Aspose.Cells | retrieve formula strings from all cells after workbook.CalculateFormula in C# | log Excel formula relationships for troubleshooting with Aspose.Cells
// Tags: Aspose.Cells workbook.CalculateFormula debugging | C# enumerate cell formulas Aspose.Cells | log Excel formula relationships .NET | trace formula graph Aspose.Cells | extract cell formula strings after calculation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, runs Workbook.CalculateFormula to evaluate all formulas, iterates through every worksheet and cell to output each non‑empty formula, and then saves the workbook, providing error handling for missing files and runtime exceptions. This pattern helps developers extract and log the formula dependency graph for debugging purposes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the workbook
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during formula calculation: {ex.Message}");
            }

            // Iterate through each worksheet and its cells containing formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Console.WriteLine($"Cell {sheet.Name}!{cell.Name} Formula: {cell.Formula}");
                    }
                }
            }

            // Save the workbook after calculation (optional)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
