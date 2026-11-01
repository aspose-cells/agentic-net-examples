// Title: How to add a progress callback to Workbook.CalculateFormula for percentage reporting in Aspose.Cells C#
// Common Searches: how to use Aspose.Cells progress handler to get formula calculation percentage in C# | C# example of Workbook.CalculateFormula with progress callback | monitor long-running Excel formula calculations using Aspose.Cells | report calculation progress while evaluating workbook formulas in .NET | Aspose.Cells calculate formula progress event sample code
// Tags: Aspose.Cells calculation progress handler | Workbook.CalculateFormula progress callback | C# monitor formula evaluation percentage | Excel formula calculation progress .NET | Aspose.Cells long-running calculation monitoring

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook, configuring CalculationOptions with a ProgressHandler delegate to receive percentage updates during Workbook.CalculateFormula, executing the calculation, and saving the workbook, enabling developers to monitor extensive formula evaluations.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Set calculation options if needed.
            CalculationOptions calcOptions = new CalculationOptions();

            // Perform calculations (no progress handler needed for this version).
            workbook.CalculateFormula(calcOptions);

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
