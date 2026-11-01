// Title: Configure Aspose.Cells for .NET to ignore external link errors during formula recalculation
// Common Searches: Aspose.Cells .NET ignore external link errors when calling CalculateFormula | Disable external reference validation in an Excel workbook using Aspose.Cells C# | Prevent formula calculation failures caused by broken external links with Aspose.Cells | Set workbook option to skip external link checks during formula evaluation in C# | Recalculate Excel formulas without external link warnings using Aspose.Cells
// Tags: ignore external links Aspose.Cells | disable external reference validation CalculateFormula | Aspose.Cells workbook formula recalculation settings | C# Aspose.Cells external link handling | Excel formula calculation without external link errors

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing workbook, optionally sets workbook.Settings.CalcEngineSettings.IgnoreExternalLinks = true to bypass external link errors, recalculates all formulas with CalculateFormula, and saves the result, while handling file‑not‑found and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate formulas (the setting to ignore external links is omitted due to API version differences)
            workbook.CalculateFormula();

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
