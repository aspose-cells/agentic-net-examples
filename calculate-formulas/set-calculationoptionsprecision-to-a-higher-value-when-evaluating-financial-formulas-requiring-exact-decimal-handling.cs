// Title: Recalculate Excel formulas in C# with Aspose.Cells while ensuring high decimal precision for financial calculations
// Common Searches: how to recalculate Excel formulas with Aspose.Cells in C# for financial precision | c# Aspose.Cells exact decimal handling when evaluating workbook formulas | increase calculation accuracy for financial models using Aspose.Cells CalculateFormula | Aspose.Cells calculation precision limitations and workarounds in .NET | recalculate workbook formulas without explicit precision settings in Aspose.Cells
// Tags: Aspose.Cells recalculate formulas C# | Excel workbook calculation precision .NET | financial formula decimal accuracy Aspose.Cells | Workbook.CalculateFormula high precision handling | Aspose.Cells calculation options limitation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, notes that calculation precision is managed internally (no explicit precision API), recalculates all formulas using Workbook.CalculateFormula, and saves the updated file, handling missing input files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Note: In the current Aspose.Cells version, calculation precision
            // is managed internally; explicit CalcEngineSettings are not available.
            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
