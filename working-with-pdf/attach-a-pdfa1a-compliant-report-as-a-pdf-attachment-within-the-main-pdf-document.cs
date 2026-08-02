// Title: Embed a PDF/A‑1a Report as an Attachment in a PDF Generated with Aspose.Cells (C#)
// Description: Demonstrates how to create a secondary workbook, save it as a PDF/A‑1a file, add the PDF as an OLE object to a primary workbook, and export the combined workbook as a PDF/A‑1a document. The example ensures PDF/A‑1a compliance, automatically embeds the attachment, and removes the temporary file after conversion.
// Keywords: Aspose.Cells PDF/A-1a | C# embed PDF attachment | Excel OLE object PDF | PDF/A-1a compliance Aspose | add PDF attachment Aspose.Cells | generate PDF/A-1a from workbook | Aspose.Cells OLE object example
// Common Searches: How to embed a PDF/A-1a file in a PDF using Aspose.Cells C# | Aspose.Cells add OLE object PDF attachment while saving to PDF/A-1a | Create PDF/A-1a compliant document with embedded PDF in Aspose.Cells | C# Aspose.Cells attach secondary PDF to main PDF | Export Excel workbook to PDF/A-1a with attached report
// Developer Intent: Add a PDF/A‑1a report as an embedded attachment inside a PDF/A‑1a document generated from an Aspose.Cells workbook.
// Use Cases: Generate a regulatory report as PDF/A‑1a and bundle it with a master Excel‑derived PDF for long‑term archiving. | Create a primary Excel file that includes supporting documentation PDFs, then export the whole package as a single PDF/A‑1a file. | Automate production of a PDF package where each section is a separate PDF/A‑1a file attached to a main PDF using Aspose.Cells.
// AI Prompts: Write C# code with Aspose.Cells to embed an existing PDF/A‑1a file as an OLE object and save the workbook as a PDF/A‑1a document. | Explain which PdfSaveOptions settings are required to keep PDF/A‑1a compliance when attaching a PDF to another PDF in Aspose.Cells. | Show how to clean up temporary files after embedding a PDF/A‑1a report in a generated PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAttachmentDemo
{
    // Demonstrates how to create a secondary workbook, save it as a PDF/A‑1a file, add the PDF as an OLE object to a primary workbook, and export the combined workbook as a PDF/A‑1a document. The example ensures PDF/A‑1a compliance, automatically embeds the attachment, and removes the temporary file after conversion.
    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a secondary workbook that will serve as the PDF/A‑1a report
                // -----------------------------------------------------------------
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Cells["A1"].PutValue("PDF/A‑1a Compliant Report");
                reportSheet.Cells["A2"].PutValue(DateTime.Now.ToString("yyyy-MM-dd"));
                reportSheet.Cells["A3"].PutValue("Sample data row 1");
                reportSheet.Cells["A4"].PutValue("Sample data row 2");

                // Save the secondary workbook as PDF/A‑1a
                string reportPdfPath = "Report.pdf";
                PdfSaveOptions reportOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1a // PDF/A‑1a compliance
                    // EmbedAttachments must remain false for PDF/A compliance
                };
                reportWorkbook.Save(reportPdfPath, reportOptions);

                // -----------------------------------------------------------------
                // 2. Create the main workbook where the PDF report will be attached
                // -----------------------------------------------------------------
                Workbook mainWorkbook = new Workbook();
                Worksheet mainSheet = mainWorkbook.Worksheets[0];
                mainSheet.Cells["A1"].PutValue("Main Document");
                mainSheet.Cells["A2"].PutValue("The PDF report is attached as an OLE object.");

                // Ensure the report PDF exists before reading
                if (!File.Exists(reportPdfPath))
                    throw new FileNotFoundException($"Report PDF not found at '{reportPdfPath}'.");

                // Add the PDF report as an OLE object (attachment)
                byte[] pdfBytes = File.ReadAllBytes(reportPdfPath);
                int oleIndex = mainSheet.OleObjects.Add(5, 0, 200, 200, pdfBytes);
                // Optional: configure the OLE object if needed (requires OleObject type)
                // var ole = mainSheet.OleObjects[oleIndex];
                // ole.FileFormatType = FileFormatType.Pdf;
                // ole.DisplayAsIcon = true;

                // -----------------------------------------------------------------
                // 3. Save the main workbook as PDF with PDF/A‑1a compliance
                // -----------------------------------------------------------------
                PdfSaveOptions mainOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1a // PDF/A‑1a compliance for the output PDF
                    // EmbedAttachments must stay false when Compliance is set
                    // The OLE object will be embedded automatically as an attachment
                };
                string outputPdfPath = "MainDocumentWithAttachment.pdf";
                mainWorkbook.Save(outputPdfPath, mainOptions);

                // Clean up temporary report file
                if (File.Exists(reportPdfPath))
                    File.Delete(reportPdfPath);

                Console.WriteLine($"Main PDF saved to '{outputPdfPath}' with the PDF/A‑1a report attached.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
