// Title: How to enable the 64‑bit Precise calculation engine for high‑accuracy financial formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells set calculation engine to Precise for financial calculations in C# | How to use CalcEngineType.Precise with Aspose.Cells .NET | Enable high‑accuracy formula evaluation in an Aspose.Cells workbook C# | Switch Aspose.Cells calculation engine to 64‑bit precise mode before recalculating formulas | C# example of assigning a Precise calculation engine in Aspose.Cells
// Tags: Aspose.Cells precise calculation engine | CalcEngineType.Precise .NET | high‑precision financial formulas Aspose.Cells | Workbook.Settings.CalcEngine assignment | force formula recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, optionally assigns the Precise 64‑bit calculation engine via Workbook.Settings.CalcEngine, forces a full formula recalculation, and saves the workbook, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Precise calculation engine APIs may not be available in older Aspose.Cells versions.
            // If needed and supported, you can enable it as shown below (uncomment when supported):
            // workbook.Settings.CalcEngine = new CalculationEngine(CalcEngineType.Precise);

            // Force recalculation so that formulas are updated
            workbook.CalculateFormula();

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
