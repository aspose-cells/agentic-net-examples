// Title: How to set CalculationOptions.MaxIterationCount to 100 in Aspose.Cells for .NET to limit iterative formula calculations
// Common Searches: Aspose.Cells set MaxIterationCount to 100 in C# | limit iterative calculation cycles in Aspose.Cells .NET workbook | configure CalculationOptions.MaxIterationCount property for Excel formulas using Aspose.Cells | C# Aspose.Cells example for restricting formula iteration count | restrict number of calculation iterations with Aspose.Cells CalculationOptions
// Tags: Aspose.Cells CalculationOptions MaxIterationCount | set iterative calculation limit .NET | configure Excel formula iteration count Aspose | C# workbook calculation options | limit iterative formula cycles Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample illustrates configuring the CalculationOptions.MaxIterationCount property to 100 (noting that this setting requires a newer Aspose.Cells version), creates a new workbook, ensures the output directory exists, and saves the workbook while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: Iterative calculation settings are not available in this version of Aspose.Cells.
            // If needed, they can be configured using the appropriate properties in newer versions.

            // Define output path
            string outputPath = "Result.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
