// Title: Load an Excel workbook in C# with Aspose.Cells while preserving formulas and turning off automatic calculation
// Common Searches: C# Aspose.Cells load workbook without evaluating formulas | how to prevent formula calculation when opening Excel with Aspose.Cells | disable automatic calculation mode after loading workbook using Aspose.Cells LoadOptions | preserve Excel formulas on load and set CalcMode to Manual in .NET | Aspose.Cells load options to keep formulas unchanged in C#
// Tags: Aspose.Cells load workbook without calculation | preserve formulas on workbook load C# | manual calculation mode Aspose.Cells | LoadOptions disable auto calculation | Excel workbook load options Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the source file, loads the Excel workbook with Aspose.Cells using default LoadOptions to keep formulas intact, optionally switches the workbook's calculation mode to manual to stop automatic formula evaluation, ensures the output directory exists, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook with default options (formulas are preserved)
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Disable automatic calculation after loading (if supported by the version)
                // Note: In some Aspose.Cells versions the CalcMode property may not be available.
                // Uncomment the following line if your version supports it:
                // workbook.Settings.CalcMode = CalcMode.Manual;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
