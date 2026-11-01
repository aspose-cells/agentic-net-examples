// Title: Remove all empty rows from the first worksheet of an XLSX file using Aspose.Cells for .NET with DeleteOptions.UpdateReference enabled
// AI Prompts: Generate C# code that opens an XLSX workbook with Aspose.Cells, deletes blank rows on the first sheet while keeping formulas updated, and saves the result. | Show how to configure DeleteOptions.UpdateReference = true and pass it to Cells.DeleteBlankRows in a .NET application. | Create a robust C# example that checks for the input file, removes empty rows from the first worksheet, handles exceptions, and writes the output file.
// Common Searches: Aspose.Cells C# delete blank rows first worksheet update references | How to keep formulas intact when removing empty rows with Aspose.Cells .NET | DeleteBlankRows with DeleteOptions.UpdateReference true example in C# | Remove empty rows from an Excel sheet using Aspose.Cells and preserve cell references | C# code to delete blank rows in the first worksheet of an XLSX file using Aspose.Cells
// Tags: Aspose.Cells DeleteBlankRows with DeleteOptions | update references after row deletion .NET | remove empty rows first worksheet XLSX | preserve formulas Aspose.Cells DeleteBlankRows | C# Aspose.Cells blank row removal

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads input.xlsx, deletes all blank rows on the first worksheet with DeleteOptions.UpdateReference set to true to keep formulas and references correct, saves the modified workbook as output.xlsx, and includes file‑existence checking and exception handling.
    class Program
    {
        static void Main(string[] args)
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

            try
            {
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (index 0)
                Worksheet firstSheet = workbook.Worksheets[0];

                // Configure delete options to update references after deletion
                DeleteOptions deleteOptions = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank rows on the first worksheet using the specified options
                firstSheet.Cells.DeleteBlankRows(deleteOptions);

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
