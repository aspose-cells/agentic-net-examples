// Title: Enable multi‑threaded formula calculation in Aspose.Cells for large workbooks using C#
// Common Searches: c# how to turn on multi‑threaded calculation in Aspose.Cells | set calculation thread count Aspose.Cells workbook .NET | speed up formula evaluation in large Excel files with Aspose.Cells | use reflection to access EnableMultiThreadedCalculation in older Aspose.Cells versions
// Tags: Aspose.Cells multi-threaded calculation | multi-threaded calculation flag Aspose.Cells | NumberOfThreads property Aspose.Cells | parallel formula recalculation .NET | large workbook performance optimization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, uses reflection to enable the multi‑threaded calculation flag and set the number of calculation threads when supported, recalculates all formulas, and saves the updated file, improving performance for large workbooks.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to enable multi‑threaded calculation if the properties exist in the current Aspose.Cells version
            var settings = workbook.Settings;
            var mtProp = settings.GetType().GetProperty("EnableMultiThreadedCalculation");
            if (mtProp != null && mtProp.CanWrite)
            {
                mtProp.SetValue(settings, true);
            }

            var threadsProp = settings.GetType().GetProperty("NumberOfThreads");
            if (threadsProp != null && threadsProp.CanWrite)
            {
                threadsProp.SetValue(settings, 4); // Adjust thread count as needed
            }

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
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
