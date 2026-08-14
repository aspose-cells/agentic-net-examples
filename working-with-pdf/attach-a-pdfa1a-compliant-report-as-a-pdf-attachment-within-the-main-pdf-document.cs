// Title: Embed a PDF/A‑1a Report as an Attachment in a PDF using Aspose.Cells for .NET (C#)
// Description: Create an Excel workbook, generate a minimal PDF/A‑1a file, add it as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the workbook so the PDF/A‑1a report is embedded inside the resulting PDF.
// Keywords: Aspose.Cells | C# | PDF/A-1a | embed PDF attachment | OLE object | PdfSaveOptions | export Excel to PDF | PDF attachment Aspose.Cells | PDF/A compliance | Aspose.Cells PDF options
// Common Searches: Aspose.Cells embed PDF/A-1a attachment | How to add OLE object PDF in Excel with Aspose.Cells | PdfSaveOptions EmbedAttachments example C# | Export workbook to PDF with embedded files | Attach PDF file to generated PDF using Aspose.Cells
// Developer Intent: Add a PDF/A‑1a report to an Excel workbook as an OLE object and ensure it is embedded in the PDF produced by Aspose.Cells.
// Use Cases: Include a certified audit PDF/A‑1a report with a financial Excel summary for regulator review. | Attach product specification PDFs to a catalog created in Excel and deliver a single PDF package to customers. | Automate compliance documentation by bundling policy PDFs with generated Excel‑to‑PDF reports.
// AI Prompts: Show C# code that adds a PDF/A‑1a file as an OLE object in an Aspose.Cells worksheet and embeds it when saving to PDF. | Explain the required PdfSaveOptions settings to include OLE attachments in the output PDF. | Provide a step‑by‑step example of creating a temporary PDF/A‑1a file, attaching it to a workbook, and exporting a single PDF with the attachment.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAttachmentDemo
{
    // Create an Excel workbook, generate a minimal PDF/A‑1a file, add it as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the workbook so the PDF/A‑1a report is embedded inside the resulting PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a sample workbook with some content
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Main PDF Document");
                sheet.Cells["A2"].PutValue("See attached PDF/A‑1a report.");

                // -----------------------------------------------------------------
                // 2. Prepare a PDF/A‑1a compliant report to be attached
                //    (For demo purposes we just create a simple PDF file.)
                // -----------------------------------------------------------------
                string attachedPdfPath = "Report_PdfA1a.pdf";

                // Create a minimal PDF file (self‑contained example)
                byte[] pdfBytes = Encoding.ASCII.GetBytes(
                    "%PDF-1.4\n%âãÏÓ\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
                    "2 0 obj\n<< /Type /Pages /Count 0 >>\nendobj\nxref\n0 3\n" +
                    "0000000000 65535 f \n0000000010 00000 n \n0000000060 00000 n \n" +
                    "trailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n115\n%%EOF");
                try
                {
                    File.WriteAllBytes(attachedPdfPath, pdfBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write temporary PDF file: {ex.Message}");
                    return;
                }

                // -----------------------------------------------------------------
                // 3. Embed the PDF file as an OLE object in the worksheet
                // -----------------------------------------------------------------
                if (File.Exists(attachedPdfPath))
                {
                    // Add the OLE object at row 5, column 1 (zero‑based indices)
                    int oleIndex = sheet.OleObjects.Add(5, 1, 200, 200, File.ReadAllBytes(attachedPdfPath));

                    // Optional: configure the OLE object if the API version supports it
                    // (Properties such as FileFormatType, DisplayAsIcon, IsObjectLink may be set here.)
                }
                else
                {
                    Console.WriteLine($"Attachment file not found: {attachedPdfPath}");
                    return;
                }

                // -----------------------------------------------------------------
                // 4. Configure PDF save options to embed OLE attachments
                // -----------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // -----------------------------------------------------------------
                // 5. Save the workbook as a PDF file; the attached PDF will be embedded
                // -----------------------------------------------------------------
                string outputPdf = "MainDocument_WithAttachment.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up the temporary attached PDF file
                try
                {
                    if (File.Exists(attachedPdfPath))
                    {
                        File.Delete(attachedPdfPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary PDF file: {ex.Message}");
                }

                Console.WriteLine($"PDF generated: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
