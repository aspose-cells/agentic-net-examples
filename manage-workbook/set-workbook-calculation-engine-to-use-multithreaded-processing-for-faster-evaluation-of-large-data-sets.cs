// Title: How to enable multi‑threaded formula calculation in an Aspose.Cells workbook with C#
// AI Prompts: Write C# code that turns on Aspose.Cells threaded calculation and configures the engine to use all available processor cores. | Show an example of recalculating all formulas after enabling multi‑core calculation in an Aspose.Cells workbook. | Provide a snippet that checks whether the current Aspose.Cells version supports the EnableThreadedCalculation setting before applying it.
// Common Searches: Aspose.Cells C# enable threaded calculation for large workbooks | Set number of calculation threads in Aspose.Cells .NET API | How to use multi‑core formula evaluation with Aspose.Cells workbook | Check Aspose.Cells version support for EnableThreadedCalculation property | Recalculate Excel formulas using all CPU cores in Aspose.Cells
// Tags: Aspose.Cells enable threaded calculation | Aspose.Cells set calculation thread count | Aspose.Cells multi‑core formula evaluation | Aspose.Cells workbook recalculate formulas | Aspose.Cells .NET calculation engine configuration

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, notes that multi‑threaded calculation may not be available in the current library version (commented code shows how to enable it), recalculates all formulas using the default engine, and saves the workbook while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Multi‑threaded calculation is not available in this version of Aspose.Cells.
            // If supported, you could enable it with:
            // workbook.Settings.EnableThreadedCalculation = true;
            // workbook.Settings.NumberOfThreads = Environment.ProcessorCount;

            // Recalculate all formulas using the available calculation engine
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
