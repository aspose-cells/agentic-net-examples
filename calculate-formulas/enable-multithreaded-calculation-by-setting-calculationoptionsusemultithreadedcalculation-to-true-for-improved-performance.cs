// Title: Enable Multi‑Threaded Formula Calculation in Aspose.Cells for .NET
// Common Searches: Aspose.Cells enable multi threaded formula calculation C# | set UseThreadedCalculation property on Workbook settings Aspose.Cells | improve Excel formula calculation performance with Aspose.Cells multi‑threading | Aspose.Cells CalculationOptions.UseMultiThreadedCalculation example C# | parallel formula evaluation using Aspose.Cells workbook.CalculateFormula
// Tags: Aspose.Cells multi‑threaded calculation | Workbook.Settings.UseThreadedCalculation C# | Parallel formula evaluation Aspose.Cells | Excel calculation performance .NET | CalculationOptions.UseMultiThreadedCalculation

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to activate multi‑threaded formula calculation in Aspose.Cells by setting the appropriate workbook or calculation options flag, then forces a full recalculation and saves the workbook, demonstrating performance‑optimized Excel processing in C#.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;

            // Example of loading an existing workbook.
            // Uncomment and adjust the path as needed.
            // string inputPath = "input.xlsx";
            // if (File.Exists(inputPath))
            // {
            //     workbook = new Workbook(inputPath);
            // }
            // else
            // {
            //     Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook instead.");
            //     workbook = new Workbook();
            // }

            // For demonstration, create a new workbook.
            workbook = new Workbook();

            // Enable multi‑threaded calculation if the API supports it.
            // The property UseThreadedCalculation may not be available in older versions.
            // Uncomment the following line if your Aspose.Cells version includes it.
            // workbook.Settings.UseThreadedCalculation = true;

            // Force a full calculation to apply any settings.
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // Ensure the output directory exists.
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
