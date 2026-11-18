// Title: Convert an XLSB workbook to PDF in C# with Aspose.Cells using PdfSaveOptions (OnePagePerSheet)
// Common Searches: C# Aspose.Cells convert XLSB file to PDF with each worksheet on a separate page | How to use PdfSaveOptions when exporting an XLSB workbook to PDF in .NET | Aspose.Cells PDF export settings for XLSB workbooks in C# | Validate file existence before converting XLSB to PDF with Aspose.Cells | Exception handling for XLSB to PDF conversion using Aspose.Cells in C#
// Tags: XLSB to PDF conversion Aspose.Cells | PdfSaveOptions OnePagePerSheet usage | Aspose.Cells workbook Save method C# | Office Add‑Ins rendering during PDF export | C# file existence check before Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies that the input XLSB file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to place each worksheet on its own PDF page (OnePagePerSheet), and saves the workbook as a PDF file while handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsb";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the XLSB workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (using available properties)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example option: place each worksheet on a separate PDF page
                    OnePagePerSheet = true
                };

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
