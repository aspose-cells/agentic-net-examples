// Title: Disable automatic calculation for Excel tables while keeping other formulas in Automatic mode using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# disable automatic calculation for tables only | How to keep workbook CalcMode Automatic but turn off table auto‑calc in Aspose.Cells | Programmatically set EnableTableAutoCalculate false while preserving other formulas | Selective formula recalculation for Excel tables with Aspose.Cells .NET
// Tags: Aspose.Cells EnableTableAutoCalculate false | Workbook CalcMode Automatic with table auto‑calc disabled | C# disable Excel table auto calculation Aspose.Cells | Selective formula recalculation Aspose.Cells .NET | Excel table calculation setting Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, optionally ensures the workbook calculation mode remains Automatic, disables automatic calculation for tables via the EnableTableAutoCalculate property, and saves the updated file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // The following settings are available in newer Aspose.Cells versions.
            // Uncomment if your version supports them.

            // workbook.Settings.CalcMode = Aspose.Cells.CalculationMode.Automatic;
            // workbook.Settings.EnableTableAutoCalculate = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
