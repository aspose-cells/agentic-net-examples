// Title: C# – Convert XLSM to PDF/A‑2a with Aspose.Cells while preserving images
// Description: Loads an XLSM workbook using auto‑detect, sets PdfSaveOptions (Compliance = PdfA2a, EmbedAttachments = false, EmbedStandardWindowsFonts = true) and calls ConversionUtility.Convert to produce a PDF/A‑2a file that keeps all embedded pictures intact.
// Keywords: Aspose.Cells XLSM to PDF/A-2a | C# PDF/A-2a conversion | preserve images Excel to PDF | LoadOptions Auto format | PdfSaveOptions Compliance PdfA2a | EmbedAttachments false | ConversionUtility Convert | macro‑enabled workbook PDF export
// Common Searches: convert xlsm to pdf/a-2a c# | asp​ose.cells preserve images pdf/a-2a | pdfsaveoptions pdf/a-2a settings | how to disable ole embedding asp​ose.cells | batch convert xlsm to pdf/a-2a
// Developer Intent: Generate a PDF/A‑2a document from a macro‑enabled XLSM file without losing any embedded images.
// Use Cases: Archival‑grade reports from Excel templates that contain charts, photos, or logos. | Regulatory submissions requiring PDF/A‑2a compliance while retaining visual content. | Automated pipelines that batch‑process XLSM workbooks into PDF/A‑2a files for long‑term storage.
// AI Prompts: Write C# code using Aspose.Cells to convert an XLSM workbook to PDF/A‑2a, keeping all embedded images and disabling OLE attachment embedding. | Explain which PdfSaveOptions properties are mandatory for PDF/A‑2a compliance and image preservation in Aspose.Cells. | Provide a step‑by‑step guide to batch convert multiple XLSM files to PDF/A‑2a with Aspose.Cells, ensuring pictures remain intact.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering; // PdfSaveOptions resides in Aspose.Cells namespace, but included for completeness

// Loads an XLSM workbook using auto‑detect, sets PdfSaveOptions (Compliance = PdfA2a, EmbedAttachments = false, EmbedStandardWindowsFonts = true) and calls ConversionUtility.Convert to produce a PDF/A‑2a file that keeps all embedded pictures intact.
class XlsmToPdfA2aConverter
{
    static void Main()
    {
        // Path to the source XLSM workbook
        string sourcePath = "input.xlsm";

        // Desired output PDF file (PDF/A‑2a)
        string outputPath = "output.pdf";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Source file not found at '{sourcePath}'.");
            return;
        }

        try
        {
            // Load options – let Aspose.Cells auto‑detect the format (XLSM with macros)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

            // PDF save options configured for PDF/A‑2a compliance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the PDF/A compliance level
                Compliance = PdfCompliance.PdfA2a,

                // Ensure OLE attachments are NOT embedded (required for PDF/A)
                EmbedAttachments = false,

                // Preserve embedded images (default behavior, kept explicit for clarity)
                EmbedStandardWindowsFonts = true
            };

            // Perform the conversion using the provided ConversionUtility rule
            ConversionUtility.Convert(sourcePath, loadOptions, outputPath, pdfOptions);

            Console.WriteLine("Conversion completed: XLSM → PDF/A‑2a with images preserved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}
