// Title: Convert an Excel workbook with external linked images to a PDF that embeds the images using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook containing externally linked pictures, configures PdfSaveOptions to embed those images, and saves the workbook as a PDF with document structure preserved using Aspose.Cells. | Demonstrate how to enable embedding of linked images and OLE attachments during Excel‑to‑PDF conversion by setting the appropriate properties on PdfSaveOptions in a C# Aspose.Cells project.
// Common Searches: Aspose.Cells C# embed linked pictures when exporting Excel to PDF | How to include external image links in PDF generated from an .xlsx using Aspose.Cells | PdfSaveOptions EmbedAttachments true to embed images in PDF conversion Aspose.Cells | Convert Excel workbook with linked images to accessible PDF using Aspose.Cells .NET
// Tags: Aspose.Cells Excel to PDF conversion with embedded linked images | PdfSaveOptions EmbedAttachments property Aspose.Cells | preserve document structure during PDF export Aspose.Cells | C# load workbook with external image links Aspose.Cells | embed OLE attachments in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The program loads an Excel workbook that contains externally linked images, sets PdfSaveOptions to embed those images and preserve document structure, and saves the workbook as a PDF with all linked pictures embedded.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains external linked images
            string sourceExcel = "LinkedImagesWorkbook.xlsx";

            // Path where the resulting PDF will be saved
            string outputPdf = "LinkedImagesWorkbook.pdf";

            try
            {
                // Verify that the source Excel file exists to avoid FileNotFoundException
                if (!File.Exists(sourceExcel))
                {
                    Console.WriteLine($"Error: Source file not found: {sourceExcel}");
                    return;
                }

                // Load the workbook (external links will be resolved automatically)
                Workbook workbook = new Workbook(sourceExcel);

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Embed OLE attachments (if any) – this also forces linked images to be embedded
                    EmbedAttachments = true,

                    // Preserve the document structure for better accessibility
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF with embedded images: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
