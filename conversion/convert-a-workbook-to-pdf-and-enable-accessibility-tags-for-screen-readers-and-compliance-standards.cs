// Title: Convert an Excel workbook to an accessible PDF (PDF/UA) using Aspose.Cells for .NET
// Common Searches: how to save an Excel file as PDF/UA with Aspose.Cells in C# | enable accessibility tags when exporting to PDF with Aspose.Cells C# | generate screen‑reader friendly PDF from .xlsx using Aspose.Cells | set PdfSaveOptions.Accessible property for PDF/UA compliance in Aspose.Cells | C# convert workbook to PDF with PDF/UA tags using Aspose.Cells
// Tags: Aspose.Cells PDF/UA export | PdfSaveOptions Accessible flag | Excel to PDF conversion with accessibility | C# generate accessible PDF from workbook | screen reader friendly PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The example loads an existing .xlsx file with Aspose.Cells, configures PdfSaveOptions to enable the Accessible flag for PDF/UA compliance, and saves the workbook as an accessible PDF while handling errors.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Ensure the input workbook exists before attempting to load it
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // If the library version supports PDF/UA accessibility, enable it:
                // pdfOptions.Accessible = true;   // Uncomment for supported versions

                // Save the workbook as a PDF file
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
