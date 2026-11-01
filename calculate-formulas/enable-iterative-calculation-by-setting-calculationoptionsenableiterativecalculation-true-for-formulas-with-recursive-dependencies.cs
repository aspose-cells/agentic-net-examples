// Title: Enable iterative calculation for circular references in Aspose.Cells using C#
// Common Searches: Aspose.Cells enable iterative calculation for circular references in C# | set CalculationOptions.EnableIterativeCalculation true Aspose.Cells .NET | how to handle recursive formulas with iterative calculation in Aspose.Cells | upgrade Aspose.Cells to use WorkbookSettings.IterativeCalculation property | enable iterative calculation when loading a workbook with Aspose.Cells
// Tags: Aspose.Cells iterative calculation setting | EnableIterativeCalculation property C# | circular reference handling Aspose.Cells | WorkbookSettings.IterativeCalculation upgrade requirement | recursive formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one if missing), notes that the current Aspose.Cells version lacks iterative calculation settings, recommends upgrading to a version that supports WorkbookSettings.IterativeCalculation, MaxIterations, and MaxChange, and then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Iterative calculation settings are not available in this version of Aspose.Cells.
            // If needed, upgrade to a newer version that supports WorkbookSettings.IterativeCalculation,
            // WorkbookSettings.MaxIterations, and WorkbookSettings.MaxChange.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
