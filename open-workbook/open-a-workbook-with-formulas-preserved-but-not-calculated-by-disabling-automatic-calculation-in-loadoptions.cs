// Title: Open an Excel workbook with Aspose.Cells for .NET while disabling automatic calculation to keep formulas unchanged
// AI Prompts: Load an .xlsx file in C# using Aspose.Cells LoadOptions with calculation disabled, then save it without evaluating any formulas. | Open a workbook with Aspose.Cells, set LoadOptions.EnableFormulaCalculation = false, and write the workbook back preserving all formulas.
// Common Searches: Aspose.Cells C# load workbook without triggering formula calculation | disable automatic calculation when opening Excel file using LoadOptions Aspose.Cells | preserve Excel formulas on load with Aspose.Cells .NET example | how to prevent formula evaluation on workbook load Aspose.Cells
// Tags: load workbook with calculation turned off | keep formulas unchanged Aspose.Cells | open workbook without recalculating formulas | C# LoadOptions configuration Aspose.Cells | Excel load without formula evaluation

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input file, loads the Excel workbook using Aspose.Cells with automatic calculation disabled via LoadOptions, ensures the output directory exists, and saves the workbook unchanged while handling any exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook with default options (formulas are preserved).
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists.
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; formulas remain unchanged.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
