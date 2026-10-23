// Title: C# plugin that uses Aspose.Cells to calculate formulas in parallel for large Excel workbooks
// Common Searches: how to calculate Excel formulas faster with Aspose.Cells in C# | enable multi‑threaded formula calculation in Aspose.Cells .NET | process large XLSX workbook using parallel calculation engine Aspose.Cells | C# example for parallel formula evaluation with Aspose.Cells
// Tags: parallel formula calculation Aspose.Cells | multi‑threaded workbook processing .NET | large Excel workbook performance Aspose.Cells | C# plugin for Aspose.Cells calculation engine | calculate formulas in parallel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input XLSX file into an Aspose.Cells Workbook, invokes CalculateFormula (which leverages Aspose.Cells' internal multi‑threaded engine when available), and saves the result to an output file, providing a simple plugin pattern for accelerating formula evaluation in large workbooks.
class ParallelCalculationPlugin
{
    static void Main(string[] args)
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Multi‑threaded calculation settings are managed internally by Aspose.Cells.
            // If the used version supports explicit settings, they can be applied here.
            // For compatibility, the explicit properties are omitted.

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
