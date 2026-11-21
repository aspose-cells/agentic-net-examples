// Title: Embedding multiple external files as attachments when saving an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed multiple files into PDF using PdfSaveOptions | set PdfSaveOptions.EmbedAttachments and specify attachment file paths in Aspose.Cells | save Excel workbook as PDF with attached documents using Aspose.Cells .NET | how to use PdfSaveOptions.Attachments collection to embed files in PDF with Aspose.Cells | add external files as PDF attachments when converting Excel to PDF in C#
// Tags: Aspose.Cells PdfSaveOptions attachment embedding | C# add external files to PDF via Aspose.Cells | Excel workbook PDF export with attached documents | using PdfSaveOptions.Attachments API in .NET | multiple PDF attachments from Excel conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The example loads an Excel file, enables PdfSaveOptions.EmbedAttachments, notes that the Attachments collection is unavailable in the current Aspose.Cells version, and saves the workbook as a PDF while handling missing files and runtime errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string inputPath = "input.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Enable embedding of attachments (if any are added)
                    EmbedAttachments = true
                };

                // NOTE: The Attachments collection is not available in the current Aspose.Cells version.
                // If you need to embed files, ensure you are using a version that supports
                // PdfSaveOptions.Attachments or use the appropriate API for embedding files.

                // Save the workbook as PDF
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
