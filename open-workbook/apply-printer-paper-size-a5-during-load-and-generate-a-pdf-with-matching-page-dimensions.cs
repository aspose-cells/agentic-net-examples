// Title: Apply A5 paper size to all worksheets while loading an Excel file and export to PDF with Aspose.Cells for .NET
// Common Searches: how to set A5 paper size for each worksheet in Aspose.Cells C# | export Excel workbook to PDF with A5 page dimensions using Aspose.Cells | Aspose.Cells C# set PageSetup PaperSize before saving as PDF | C# load workbook and change printer paper size to A5 for PDF output | Aspose.Cells PDF conversion preserving custom page size A5
// Tags: worksheet page setup A5 Aspose.Cells | PDF export with specific paper size C# | apply printer paper size to all worksheets | Aspose.Cells save workbook as PDF custom dimensions | load Excel workbook set paper size programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, iterates through every worksheet to set its PageSetup.PaperSize to A5, ensures the output folder exists, and then saves the workbook as a PDF, handling both Aspose.Cells‑specific and general exceptions.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(inputPath);

                // Set paper size to A5 for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handle Aspose.Cells specific errors (e.g., password protected files)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle all other runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
