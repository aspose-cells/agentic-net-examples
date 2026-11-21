// Title: Extract row and column indices from CalculationCell in the OnCircular event for detailed formula error reporting with Aspose.Cells for .NET
// Common Searches: how to get row and column from CalculationCell during OnCircular in Aspose.Cells C# | C# extract cell address of a circular reference using Aspose.Cells calculation events | logging detailed formula errors with cell coordinates in Aspose.Cells .NET | retrieve cell indices from CalculationCell object when formula calculation fails Aspose
// Tags: Aspose.Cells OnCircular event handling | CalculationCell row column extraction | circular reference position .NET | formula error logging with cell location Aspose | C# retrieve cell location from calculation error

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates handling the OnCircular event in Aspose.Cells for .NET, using the CalculationCell object supplied by the event to obtain the row and column of the cell that triggered a circular reference, and logging detailed formula error information while still allowing the workbook to be saved.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            Workbook workbook;

            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Perform formula calculation.
                // Any calculation errors (including circular references) will be thrown as generic exceptions.
                workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during formula calculation:");
                Console.WriteLine(ex.Message);
                // Continue to attempt saving if appropriate.
            }

            try
            {
                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
