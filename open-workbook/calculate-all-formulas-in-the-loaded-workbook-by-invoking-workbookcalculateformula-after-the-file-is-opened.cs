// Title: Calculate all formulas in an Excel workbook using Aspose.Cells for .NET (C#) and save the updated file
// Common Searches: C# Aspose.Cells calculate all workbook formulas before saving | How to force formula recalculation with Aspose.Cells Workbook.CalculateFormula | Programmatically evaluate Excel formulas in .NET using Aspose.Cells | Aspose.Cells example to recalculate formulas after loading a workbook | Run Workbook.CalculateFormula in C# and export the updated Excel file
// Tags: Aspose.Cells Workbook.CalculateFormula C# | recalculate Excel formulas .NET | load workbook with LoadOptions Aspose.Cells | save workbook after formula calculation | handle CellsException during workbook processing

using System;
using System.IO;
using Aspose.Cells;

// The C# program checks for the input file, loads it into an Aspose.Cells Workbook (optionally using LoadOptions), invokes Workbook.CalculateFormula() to evaluate every formula, and saves the result to a new file while handling Aspose.Cells‑specific and generic exceptions.
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
            // LoadOptions can be used to specify a password if required.
            // Here we leave the password empty; if the workbook is protected,
            // the catch block will handle the exception.
            var loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            var workbook = new Workbook(inputPath, loadOptions);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password required)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other runtime errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
