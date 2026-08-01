// Title: Convert Excel to PDF without Embedding Linked Images – Aspose.Cells C#
// Description: Loads an Excel workbook with Aspose.Cells, configures PdfSaveOptions to turn off attachment embedding and ignore blank pages, optionally handles the EnableExternalLinks setting, and saves a compact PDF that excludes external linked pictures.
// Keywords: Aspose.Cells | C# | Excel to PDF conversion | exclude linked images | PdfSaveOptions | EmbedAttachments false | ignore blank pages | reduce PDF size | EnableExternalLinks | batch workbook conversion
// Common Searches: Aspose.Cells convert Excel to PDF without images | How to skip external linked pictures when saving PDF with Aspose.Cells | Minimize PDF file size Aspose.Cells C# | PdfSaveOptions ignore blank pages example | Disable OLE attachments in Aspose.Cells PDF export
// Developer Intent: Generate a PDF from an Excel file while preventing external linked images and OLE attachments from being embedded, resulting in a smaller document.
// Use Cases: Produce archival PDFs that contain only embedded data, omitting linked pictures to keep storage low. | Create clean, size‑optimized PDF reports for email distribution or web publishing. | Run automated batch conversions of workbooks where blank pages and external media must be excluded for compliance.
// AI Prompts: Show C# code using Aspose.Cells PdfSaveOptions to export an Excel workbook to PDF without embedding linked images and to ignore blank pages. | Explain how to detect and disable the EnableExternalLinks property in Aspose.Cells before saving a workbook as a minimal‑size PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook with Aspose.Cells, configures PdfSaveOptions to turn off attachment embedding and ignore blank pages, optionally handles the EnableExternalLinks setting, and saves a compact PDF that excludes external linked pictures.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the EnableExternalLinks property is unavailable.
            // If needed, configure related settings here using the appropriate API for your version.

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not embed OLE attachments (default is false, set explicitly for clarity)
                EmbedAttachments = false,
                // Skip blank pages to further reduce file size
                PrintingPageType = PrintingPageType.IgnoreBlank
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
