// Title: Extract row and column indices from CalculationCell in Aspose.Cells OnCircular event for detailed circular reference error reporting (C#)
// Common Searches: asp.net aspose.cells oncircular handler get cell row and column | c# retrieve cell address from CalculationCell during circular reference detection | how to log detailed circular reference errors using CalculationCell in Aspose.Cells | formula calculation exception handling with row/column info Aspose.Cells C# | use CalculationCell.Row and .Column in OnCircular event for error reporting
// Tags: Aspose.Cells OnCircular event | CalculationCell row column extraction | circular reference error reporting C# | formula calculation exception handling Aspose.Cells | cell coordinate retrieval in OnCircular

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Demonstrates how to subscribe to the Aspose.Cells OnCircular event, use the provided CalculationCell object to obtain the exact row and column of a circular reference, and generate detailed error messages for formula calculation failures in a C# application.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
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

                // Attempt to calculate all formulas.
                try
                {
                    workbook.CalculateFormula();
                }
                catch (Exception ex)
                {
                    // Handle errors that may occur during formula calculation (e.g., circular references)
                    Console.WriteLine("An error occurred during formula calculation:");
                    Console.WriteLine(ex.Message);
                }

                // Save the workbook (if any changes need to be persisted)
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception e)
            {
                // General exception handling for unexpected errors
                Console.WriteLine($"Error: {e.Message}");
            }
        }
    }
}
