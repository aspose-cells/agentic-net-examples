// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET (C#) with robust file checks
// Common Searches: how to save an .xlsx workbook as PDF using Aspose.Cells in C# | Aspose.Cells C# example for converting Excel to PDF with custom save options | C# code to verify Excel file exists before converting to PDF with Aspose.Cells | Aspose.Cells PDFSaveOptions properties list for .NET | why Aspose.Cells PdfSaveOptions does not include a document language setting
// Tags: Aspose.Cells convert workbook to PDF C# | PdfSaveOptions configuration Aspose.Cells | Excel file existence validation C# | error handling Aspose.Cells conversion | document language unsupported Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing XLSX file into an Aspose.Cells Workbook, checks that the source file is present, configures PdfSaveOptions, and saves the workbook as a PDF while handling any runtime exceptions. It also notes that PdfSaveOptions does not provide a DocumentLanguage property for setting the PDF language.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no DocumentLanguage property in PdfSaveOptions)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF
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
}
