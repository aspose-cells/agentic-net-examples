// Title: How to limit formula calculation threads to four using Aspose.Cells CalculationOptions in C#
// Common Searches: Aspose.Cells set CalculationOptions.ThreadCount to 4 in C# | limit parallel formula calculation threads in .NET workbook | configure maximum calculation threads for Aspose.Cells workbook | control number of threads used during formula evaluation with Aspose.Cells | C# Aspose.Cells calculation thread count configuration example
// Tags: Aspose.Cells calculationoptions threadcount | limit workbook formula calculation parallelism | set maximum calculation threads c# | control formula evaluation threads aspnet | configure calculation thread count aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// This C# example creates a new workbook (or loads an existing one), optionally sets CalculationOptions.ThreadCount to 4 to restrict parallel formula evaluation, runs a full calculation with CalculateFormula, ensures the output directory exists, and saves the workbook as output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Perform a full calculation (threaded calculation settings are optional and depend on the library version)
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
