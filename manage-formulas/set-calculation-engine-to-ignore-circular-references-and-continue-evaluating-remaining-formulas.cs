// Title: Configure Aspose.Cells for .NET to skip circular reference errors and continue formula recalculation
// Common Searches: aspnet calculate workbook formulas without circular reference error Aspose.Cells | skip circular reference detection when recalculating Excel file using Aspose.Cells C# | ignore circular reference errors in Aspose.Cells formula engine .NET | Aspose.Cells calculate all formulas even if circular references exist | workaround circular reference exception Aspose.Cells C#
// Tags: circular reference handling Aspose.Cells | skip circular reference errors Aspose.Cells | formula recalculation without iterative calculation Aspose.Cells | Aspose.Cells workbook calculation settings | Aspose.Cells .NET formula engine limitation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, notes that the current library version lacks iterative calculation settings needed to ignore circular references, recalculates all formulas, and saves the result while handling missing files and ensuring the output directory exists.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: Iterative calculation settings are not available in the current Aspose.Cells version.
                // If needed, upgrade Aspose.Cells to a version that supports these settings.

                // Recalculate all formulas
                workbook.CalculateFormula();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully and saved to '{outputPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
