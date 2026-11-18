// Title: Convert an XLSB workbook to PDF using Aspose.Cells for .NET (C#) with default PdfSaveOptions
// Common Searches: Aspose.Cells C# example to export XLSB file to PDF | How to use PdfSaveOptions when saving an XLSB workbook as PDF in .NET | Load binary Excel workbook and convert to PDF with Aspose.Cells | Check if XLSB file exists before converting to PDF in C# | Default PDF export settings for Excel workbooks using Aspose.Cells
// Tags: XLSB to PDF conversion Aspose.Cells | PdfSaveOptions default configuration .NET | C# load binary Excel workbook | Excel file existence validation C# | Aspose.Cells workbook.Save PDF method

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading an XLSB workbook with Aspose.Cells, optionally verifying the source file, applying the default PdfSaveOptions, and saving the workbook as a PDF in C#, with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsb";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
