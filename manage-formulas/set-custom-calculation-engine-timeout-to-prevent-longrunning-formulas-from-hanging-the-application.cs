// Title: Configure a custom calculation engine timeout in Aspose.Cells to avoid hanging during formula recalculation (C#)
// Common Searches: Aspose.Cells how to limit formula calculation time in .NET | set timeout for workbook.CalculateFormula to prevent hangs in C# | prevent long running Excel formulas from freezing Aspose.Cells application | custom calculation engine timeout setting for Aspose.Cells workbook | C# Aspose.Cells example to stop infinite formula loops with timeout
// Tags: Aspose.Cells calculation timeout configuration | C# limit formula evaluation time | prevent workbook.CalculateFormula hangs | custom formula engine timeout Aspose.Cells | manage long-running Excel formulas .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, optionally configures a custom calculation engine timeout to stop long‑running formulas, recalculates all formulas with CalculateFormula, saves the updated workbook, and includes error handling for file and runtime issues.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed successfully. Saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
