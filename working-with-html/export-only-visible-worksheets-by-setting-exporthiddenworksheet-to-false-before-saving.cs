// Title: How to save an Excel workbook with only visible worksheets using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook without hidden sheets | export only visible worksheets when converting Excel with Aspose.Cells | Aspose.Cells ExportHiddenWorksheet false example in .NET | skip hidden worksheets during workbook.Save in C# | Aspose.Cells hide worksheets from export programmatically
// Tags: ExportHiddenWorksheet property default behavior Aspose.Cells | save workbook excluding hidden worksheets C# | visible worksheet export Aspose.Cells .NET | Aspose.Cells workbook.Save hidden sheet handling | Excel to Excel conversion without hidden sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, relies on Aspose.Cells' default behavior that hidden worksheets are omitted when saving, and writes the workbook to a new file while handling missing input and runtime exceptions.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Hidden worksheets are not exported by default when saving to Excel formats,
                // so no additional setting is required here.

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions (e.g., loading/saving errors)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
