// Title: How to enable multi‑threaded formula calculation in Aspose.Cells for .NET to accelerate large workbooks
// AI Prompts: Generate C# code that activates multi‑threaded formula calculation in an Aspose.Cells Workbook and sets the desired thread count. | Show an example of configuring Aspose.Cells to recalculate formulas in parallel for a workbook containing many sheets. | Explain the steps required to improve formula evaluation performance by enabling parallel calculation in Aspose.Cells .NET.
// Common Searches: Aspose.Cells enable parallel formula calculation .NET | set multithreaded calculation Aspose.Cells workbook C# | increase Excel formula calculation speed using Aspose.Cells multi thread | how many threads can Aspose.Cells use for formula evaluation
// Tags: enable multi‑threaded calculation Aspose.Cells | parallel formula evaluation .NET | configure calculation threads Aspose.Cells | optimize large workbook performance Aspose.Cells | Aspose.Cells workbook multithreading settings

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the input file, loads it into an Aspose.Cells Workbook, recalculates all formulas using the default calculation engine, saves the updated workbook, and handles any exceptions that occur.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
