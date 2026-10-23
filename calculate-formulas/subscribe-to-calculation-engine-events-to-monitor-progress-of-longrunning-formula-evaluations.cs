// Title: Subscribe to Aspose.Cells CalculationEngine events to monitor long‑running formula evaluation progress in C#
// Common Searches: how to subscribe to Aspose.Cells CalculationEngine events for formula progress in C# | track long running Excel formula calculation using Aspose.Cells .NET events | C# example of handling CalculationEngine.Progress event with Aspose.Cells | monitor formula calculation status while using Workbook.CalculateFormula in Aspose.Cells | receive start and finish notifications from Aspose.Cells calculation engine in .NET
// Tags: Aspose.Cells CalculationEngine event subscription | monitor formula calculation progress C# | long running Excel formula evaluation Aspose.Cells | handle calculation start and finish notifications .NET | track calculation progress with Aspose.Cells events

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, registers handlers for the CalculationEngine.Started, CalculationEngine.Progress, and CalculationEngine.Finished events to receive real‑time updates during a potentially long‑running Workbook.CalculateFormula call, logs the progress to the console, and finally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Perform formula calculation (may be long‑running)
            try
            {
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed.");
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Calculation error: {calcEx.Message}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
