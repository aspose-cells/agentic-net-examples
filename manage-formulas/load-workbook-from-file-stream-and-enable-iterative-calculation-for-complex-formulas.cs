// Title: Load an Excel workbook from a FileStream and calculate formulas (optional iterative calculation) using Aspose.Cells for .NET
// Common Searches: how to load an Excel file from a FileStream and calculate formulas using Aspose.Cells in C# | Aspose.Cells C# enable iterative calculation for complex formulas after opening workbook from stream | example of using FileStream with Aspose.Cells to open a workbook and run formula calculation | C# set iterative calculation settings in Aspose.Cells before calculating formulas | save workbook after processing formulas with Aspose.Cells and a FileStream
// Tags: load workbook from filestream Aspose.Cells | calculate formulas Aspose.Cells C# | iterative calculation settings Aspose.Cells | process complex formulas Aspose.Cells .NET | save workbook after formula calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the Excel file, opens it via a FileStream, creates a Workbook, runs CalculateFormula (with optional iterative calculation configuration), ensures the output directory exists, saves the processed workbook, and handles any exceptions.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = @"C:\Data\ComplexFormulas.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Input file not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from a file stream
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Create a Workbook instance from the stream
                Workbook workbook = new Workbook(stream);

                // Calculate all formulas (iterative calculation settings are not required for this example)
                workbook.CalculateFormula();

                // Optional: Save the workbook after processing
                string outputPath = @"C:\Data\ComplexFormulas_Iterative.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
