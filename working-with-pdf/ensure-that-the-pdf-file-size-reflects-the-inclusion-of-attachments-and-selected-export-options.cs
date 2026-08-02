// Title: Export a Workbook to PDF with Embedded OLE Attachment and Minimum‑Size Optimization using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a temporary text file as an OLE object, configure PdfSaveOptions to embed the attachment, apply MinimumSize optimization and Flate compression, save the result as a PDF, and delete the temporary file. The final PDF size reflects both the embedded attachment and the chosen compression settings.
// Keywords: Aspose.Cells PDF export | embed OLE attachment PDF | MinimumSize PDF optimization | Flate compression Aspose.Cells | C# embed file in PDF | Aspose.Cells PdfSaveOptions | reduce PDF file size .NET | PDF with embedded attachments
// Common Searches: Aspose.Cells embed OLE object in PDF | How to reduce PDF size with Aspose.Cells | C# save workbook as PDF with attachment | PdfSaveOptions MinimumSize example | Flate compression for PDF in Aspose.Cells
// Developer Intent: Add an OLE file to a worksheet, embed it in the exported PDF, and minimize the PDF’s footprint by using Aspose.Cells compression and optimization options.
// Use Cases: Attach a terms‑and‑conditions text file to an invoice PDF while keeping the email attachment under size limits. | Distribute a financial report PDF that includes supporting spreadsheets as embedded OLE objects without bloating the file. | Publish a product manual PDF with reference documents embedded, optimized for fast download on mobile networks.
// AI Prompts: Generate C# code that uses Aspose.Cells to embed a text file as an OLE object and export the workbook to a PDF with MinimumSize optimization and Flate compression. | Explain how EmbedAttachments, OptimizationType, and PdfCompression affect the final PDF size in Aspose.Cells. | Provide best practices for cleaning up temporary files after embedding them as OLE objects in a PDF generated with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfAttachmentDemo
{
    // Demonstrates how to create a workbook, add a temporary text file as an OLE object, configure PdfSaveOptions to embed the attachment, apply MinimumSize optimization and Flate compression, save the result as a PDF, and delete the temporary file. The final PDF size reflects both the embedded attachment and the chosen compression settings.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data to the worksheet
                worksheet.Cells["A1"].PutValue("PDF with Embedded Attachments and Optimization");

                // Create a temporary file that will be embedded as an OLE object
                string tempFilePath = "sample.txt";
                File.WriteAllText(tempFilePath, "This is a sample text file to be embedded in the PDF.");

                // Add the OLE object (attachment) to the worksheet
                // Parameters: upper left row, upper left column, height, width, byte[] of the file
                int oleIndex = worksheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(tempFilePath));
                // Optional: set display as icon
                worksheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Embed the OLE attachment into the resulting PDF
                    EmbedAttachments = true,
                    // Optimize for minimum file size (more important than print quality)
                    OptimizationType = PdfOptimizationType.MinimumSize,
                    // Use Flate compression for additional size reduction
                    PdfCompression = PdfCompressionCore.Flate
                };

                // Save the workbook as a PDF with the specified options
                string outputPdf = "PdfWithEmbeddedAttachment.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up the temporary file used for embedding
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }

                Console.WriteLine($"PDF saved to '{outputPdf}' with embedded attachment and optimization.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
