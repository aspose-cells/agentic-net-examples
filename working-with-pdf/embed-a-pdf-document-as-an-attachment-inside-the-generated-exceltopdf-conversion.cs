// Title: Embed a PDF as an attachment during Excel‑to‑PDF conversion with Aspose.Cells for .NET (C#)
// Description: Creates a temporary PDF, inserts an OLE object with a custom icon, embeds the PDF data, enables EmbedAttachments in PdfSaveOptions, and saves the workbook as a PDF that carries the PDF as an attachment.
// Keywords: Aspose.Cells | C# | Excel to PDF | embed PDF attachment | OLE object | PdfSaveOptions | EmbedAttachments | PDF attachment in Excel | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells embed PDF attachment C# | Add OLE object PDF when saving Excel as PDF | PdfSaveOptions EmbedAttachments example | C# embed PDF in generated PDF using Aspose.Cells | Excel workbook to PDF with attached PDF
// Developer Intent: Add a PDF file as an embedded attachment to the PDF produced from an Excel workbook.
// Use Cases: Attach a technical specification PDF to a financial report generated from Excel. | Bundle terms‑and‑conditions PDF with an invoice PDF created from a spreadsheet. | Provide supporting documentation (PDF) alongside a data export for regulatory compliance.
// AI Prompts: Show how to embed multiple PDFs as OLE attachments when converting an Excel workbook to PDF with Aspose.Cells. | Explain how to replace the default icon for an embedded PDF OLE object in the output PDF. | Give best‑practice error handling for missing source PDFs during OleObject.SetEmbeddedObject and PdfSaveOptions usage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a temporary PDF, inserts an OLE object with a custom icon, embeds the PDF data, enables EmbedAttachments in PdfSaveOptions, and saves the workbook as a PDF that carries the PDF as an attachment.
class EmbedPdfInExcelToPdf
{
    static void Main()
    {
        try
        {
            // Ensure a sample PDF exists for embedding
            string pdfPath = "sample.pdf";
            if (!File.Exists(pdfPath))
            {
                // Minimal PDF content (sufficient for demonstration)
                File.WriteAllText(pdfPath, "%PDF-1.4\n%âãÏÓ\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF");
            }

            // Create a new workbook and add a title
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Excel to PDF with Embedded PDF Attachment");

            // Use a minimal 1x1 PNG icon (transparent) for the OLE object
            byte[] iconData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            // Add the OLE object referencing the PDF file
            int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, iconData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Embed the PDF data into the OLE object
            byte[] pdfData = File.ReadAllBytes(pdfPath);
            ole.SetEmbeddedObject(false, pdfData, pdfPath, true, "Sample PDF");
            ole.FileFormatType = FileFormatType.Pdf; // Ensure correct format

            // Configure PDF save options to embed OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with the embedded attachment
            string outputPath = "ExcelWithEmbeddedPdf.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary PDF file if it exists
            string tempPdf = "sample.pdf";
            if (File.Exists(tempPdf))
            {
                try
                {
                    File.Delete(tempPdf);
                }
                catch
                {
                    // Ignored – cleanup failure should not crash the program
                }
            }
        }
    }
}
