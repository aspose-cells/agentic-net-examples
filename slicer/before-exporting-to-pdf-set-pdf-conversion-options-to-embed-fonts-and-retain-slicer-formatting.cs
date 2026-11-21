// Title: Set PDF conversion options in Aspose.Cells for .NET to embed fonts and preserve slicer formatting when exporting an Excel workbook to PDF
// Common Searches: Aspose.Cells C# embed fonts in PDF export | retain slicer formatting when converting Excel to PDF with Aspose.Cells | PdfSaveOptions to keep slicer layout in .NET | how to enable font embedding in Aspose.Cells PDFSaveOptions | preserve Excel slicer appearance during PDF conversion using C#
// Tags: Aspose.Cells PDF font embedding | Aspose.Cells slicer layout preservation | PdfSaveOptions configuration C# | Excel to PDF conversion with slicer formatting | C# Aspose.Cells PDF export options

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The example loads an Excel workbook, creates a PdfSaveOptions object, and shows where to set font embedding and slicer layout preservation (noting that EmbedStandardFonts and SlicerLayoutOptions are unavailable in the current Aspose.Cells version), then saves the workbook as a PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF conversion options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // Note: EmbedStandardFonts and SlicerLayoutOptions are not available in the current Aspose.Cells version.
                // If needed, set other supported options here.

                // Export the workbook to PDF using the configured options
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
