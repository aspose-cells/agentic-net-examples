// Title: Use Aspose.Cells CalculateFormula in C# to evaluate all workbook formulas with a custom calculation engine
// Common Searches: C# Aspose.Cells calculate all formulas with custom calculation engine | How to trigger workbook.CalculateFormula before saving Excel file in .NET | Evaluate Excel formulas programmatically using Aspose.Cells CalculateFormula method | Load an .xlsx, recalculate formulas, and save with Aspose.Cells in C# | Register and use a custom calculation engine for Aspose.Cells workbook calculation
// Tags: Aspose.Cells CalculateFormula method | custom formula engine integration | recalculate all workbook formulas in C# | load workbook, evaluate formulas, save with Aspose.Cells | Excel formula evaluation via Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input file, loads it into an Aspose.Cells Workbook, calls CalculateFormula to evaluate every formula using any registered custom calculation engine, ensures the output directory exists, and saves the updated workbook, handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Evaluate all formulas in the workbook
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
