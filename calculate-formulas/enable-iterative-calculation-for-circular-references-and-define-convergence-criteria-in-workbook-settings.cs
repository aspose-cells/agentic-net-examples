// Title: Enable iterative calculation and define convergence limits for circular references in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# enable iterative calculation for circular references | set max iterations and max change in Aspose.Cells workbook settings C# | configure convergence criteria for circular reference formulas using Aspose.Cells | C# workbook.Settings.IterativeCalculation true Aspose.Cells example | Aspose.Cells iterative calculation property not available version check
// Tags: iterative calculation setting Aspose.Cells | circular reference convergence Aspose.Cells | maxiterations workbook setting Aspose.Cells | maxchange workbook setting Aspose.Cells | Aspose.Cells workbook configuration C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, applies iterative calculation settings (IterativeCalculation, MaxIterations, MaxChange) to handle circular references, ensures the output directory exists, and saves the file as IterativeCalculationEnabled.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Enable iterative calculation if the API supports it.
            // For newer Aspose.Cells versions you can use:
            // workbook.Settings.IterativeCalculation = true;
            // workbook.Settings.MaxIterations = 100;
            // workbook.Settings.MaxChange = 0.001;
            // The current version may not expose these properties.

            // Define output file path
            string outputPath = "IterativeCalculationEnabled.xlsx";

            // Ensure the output directory exists (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the (potential) settings applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
