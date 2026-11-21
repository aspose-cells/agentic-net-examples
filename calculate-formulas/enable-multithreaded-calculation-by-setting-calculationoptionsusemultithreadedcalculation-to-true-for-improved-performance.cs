// Title: Enable multi‑threaded formula calculation in Aspose.Cells for .NET (C#) by setting WorkbookSettings.UseThreadedCalculation
// Common Searches: C# Aspose.Cells enable multi threaded calculation for large workbooks | How to turn on multi‑threaded formula evaluation in Aspose.Cells .NET | Set WorkbookSettings.UseThreadedCalculation property in Aspose.Cells example | Improve Excel formula calculation speed with Aspose.Cells multi‑threading | Using reflection to set UseThreadedCalculation in older Aspose.Cells versions | Force full recalculation after enabling multi‑threaded calculation Aspose.Cells
// Tags: multi-threaded formula calculation Aspose.Cells | WorkbookSettings.UseThreadedCalculation C# | optimize Excel calculation performance .NET | reflection set Aspose.Cells workbook property | force CalculateFormula after settings change | enable parallel calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, uses reflection to enable WorkbookSettings.UseThreadedCalculation (when the property exists), forces a full formula recalculation with CalculateFormula, ensures the output directory exists, and saves the workbook, providing error handling throughout.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Enable multi‑threaded calculation if supported by the current Aspose.Cells version
                // (Older versions may not expose this property; safely ignore if unavailable)
                try
                {
                    // The property may not exist in some versions; use reflection to set it when possible
                    var settingsType = typeof(WorkbookSettings);
                    var prop = settingsType.GetProperty("UseThreadedCalculation");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(workbook.Settings, true);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to set threaded calculation: {ex.Message}");
                }

                // Force a full recalculation after changing the option
                workbook.CalculateFormula();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
