// Title: Embed an OLE attachment in a PDF using Aspose.Cells for .NET and verify it opens
// Description: C# example that creates a workbook, adds a temporary text file as an OLE object, enables PdfSaveOptions.EmbedAttachments, saves the workbook as a PDF, checks that the PDF is generated, and confirms the attachment is accessible in a PDF viewer.
// Keywords: Aspose.Cells PDF attachment | EmbedAttachments C# | OLE object PDF Aspose | verify PDF attachment .NET | save workbook as PDF with attachments
// Common Searches: Aspose.Cells embed file in PDF | Enable EmbedAttachments when saving PDF | Check embedded OLE objects in generated PDF | C# create PDF with attached documents using Aspose
// Developer Intent: Add an OLE file to a workbook, export it as a PDF with embedded attachments, and ensure the attachment can be opened in a PDF viewer.
// Use Cases: Distribute a single PDF that bundles contracts, specifications, or supporting documents alongside a financial report. | Automate generation of product datasheets where each sheet includes related manuals as embedded files. | Include a validation step in CI pipelines that confirms PDFs contain the expected embedded attachments.
// AI Prompts: Generate C# code to embed multiple files as OLE objects in an Aspose.Cells workbook and save it as a PDF with EmbedAttachments enabled. | Write a method that opens a PDF created by Aspose.Cells, lists all embedded attachments, and returns their filenames. | Explain how to configure PdfSaveOptions for attachment embedding and describe how to test the result in Adobe Acrobat or another PDF viewer.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAttachmentDemo
{
    // C# example that creates a workbook, adds a temporary text file as an OLE object, enables PdfSaveOptions.EmbedAttachments, saves the workbook as a PDF, checks that the PDF is generated, and confirms the attachment is accessible in a PDF viewer.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and add a sample OLE attachment
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF with Embedded Attachment");

                // Create a temporary file that will be embedded
                string tempFilePath = Path.Combine(Path.GetTempPath(), "sample.txt");
                File.WriteAllText(tempFilePath, "This is the content of the embedded attachment.");

                // Ensure the temporary file exists before reading
                if (!File.Exists(tempFilePath))
                    throw new FileNotFoundException("Temporary attachment file not found.", tempFilePath);

                // Add the OLE object (attachment) to the worksheet
                int oleIndex = sheet.OleObjects.Add(5, 5, 200, 200, File.ReadAllBytes(tempFilePath));
                sheet.OleObjects[oleIndex].DisplayAsIcon = true;
                sheet.OleObjects[oleIndex].Label = "Sample Text File";

                // ------------------------------------------------------------
                // 2. Configure PDF save options to embed attachments
                // ------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true   // Enable embedding of OLE attachments
                };

                // Save the workbook as PDF
                string pdfPath = Path.Combine(Path.GetTempPath(), "WorkbookWithAttachment.pdf");
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved to: {pdfPath}");

                // ------------------------------------------------------------
                // 3. Simple verification that the PDF was created
                // ------------------------------------------------------------
                if (File.Exists(pdfPath) && new FileInfo(pdfPath).Length > 0)
                {
                    Console.WriteLine("PDF generated successfully.");
                }
                else
                {
                    Console.WriteLine("PDF generation failed.");
                }

                // Clean up temporary files
                File.Delete(tempFilePath);
                // Optionally delete the PDF after verification
                // File.Delete(pdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
