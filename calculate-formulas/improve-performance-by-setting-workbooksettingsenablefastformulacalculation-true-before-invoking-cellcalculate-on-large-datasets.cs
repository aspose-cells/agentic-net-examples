// Title: How to enable fast formula calculation in Aspose.Cells for C# before calling CalculateFormula on large workbooks
// Common Searches: Aspose.Cells C# enable fast formula calculation for large Excel files | Improve CalculateFormula performance by setting EnableFastFormulaCalculation true | Fast formula calculation setting in Aspose.Cells .NET for big worksheets | Speed up Excel formula evaluation with Aspose.Cells workbook settings | Best practice to calculate formulas quickly using Aspose.Cells in C#
// Tags: EnableFastFormulaCalculation workbook setting | Aspose.Cells fast formula calculation | CalculateFormula performance optimization | large Excel workbook formula evaluation .NET | C# Aspose.Cells workbook calculation speed

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates loading an Excel workbook with Aspose.Cells, enabling the fast formula calculation mode via Workbook.Settings.EnableFastFormulaCalculation, calculating all formulas, and saving the workbook, resulting in a noticeable performance improvement for large datasets.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
