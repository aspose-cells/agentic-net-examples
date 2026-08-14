// Title: Embed a PDF as an OLE attachment while converting Excel to PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a PDF‑based OLE object displayed as an icon, enable the EmbedAttachments flag in PdfSaveOptions, and save the sheet as a PDF that carries the original PDF as an attachment.
// Keywords: Aspose.Cells embed PDF | C# Excel to PDF OLE attachment | PdfSaveOptions EmbedAttachments | Aspose.Cells OLE object PDF | Excel workbook PDF attachment | .NET Aspose.Cells PDF embed | C# add PDF as OLE icon
// Common Searches: Aspose.Cells add PDF OLE object C# | embed PDF attachment when saving Excel as PDF | PdfSaveOptions EmbedAttachments example | C# Excel to PDF with embedded file using Aspose | how to attach a PDF to generated PDF in Aspose.Cells
// Developer Intent: Add a PDF file as an embedded OLE object in an Excel workbook and have it appear as an attachment in the PDF produced by Aspose.Cells.
// Use Cases: Include a product‑specification PDF inside a sales‑report PDF generated from Excel. | Attach a signed contract PDF to a financial‑statement PDF so both travel together. | Showcase embedding any binary file as an OLE icon and delivering it as an attachment in the final PDF.
// AI Prompts: Write C# code that inserts a PDF as an OLE object in an Excel worksheet and saves the workbook to PDF with the PDF attached using Aspose.Cells. | Explain how to configure PdfSaveOptions to embed OLE objects as attachments in the output PDF for Aspose.Cells .NET. | Generate a sample that embeds multiple PDFs with custom icons in a workbook and exports a single PDF containing all attachments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a PDF‑based OLE object displayed as an icon, enable the EmbedAttachments flag in PdfSaveOptions, and save the sheet as a PDF that carries the original PDF as an attachment.
class EmbedPdfAttachmentDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Excel to PDF with embedded PDF attachment");

            // Path to the PDF file that will be embedded
            string pdfFilePath = "sample.pdf";

            // Create a minimal PDF file if it does not exist (for demonstration purposes)
            if (!File.Exists(pdfFilePath))
            {
                // Simple PDF header bytes; a real PDF should be used in production
                File.WriteAllBytes(pdfFilePath, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D });
            }

            // Generate a placeholder icon (PNG) for the OLE object
            byte[] iconBytes = GeneratePlaceholderIcon();

            // Add an OLE object placeholder to the worksheet using the generated icon
            int oleIndex = worksheet.OleObjects.Add(5, 0, 200, 200, iconBytes);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Read the PDF file bytes
            byte[] pdfData = File.ReadAllBytes(pdfFilePath);

            // Embed the PDF file into the OLE object (not linked, displayed as an icon)
            oleObject.SetEmbeddedObject(
                linkToFile: false,          // Do not link, embed the data
                objectData: pdfData,        // PDF file bytes
                sourceFileName: Path.GetFileName(pdfFilePath),
                displayAsIcon: true,        // Show as an icon in Excel
                label: "Embedded PDF");     // Icon label

            // Specify that the embedded object is a PDF
            oleObject.FileFormatType = FileFormatType.Pdf;

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                EmbedAttachments = true // Enable embedding of OLE attachments
            };

            // Save the workbook as PDF; the embedded PDF will be attached to the output PDF
            workbook.Save("OutputWithEmbeddedPdf.pdf", pdfSaveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary PDF file used for embedding
            string pdfFilePath = "sample.pdf";
            if (File.Exists(pdfFilePath))
            {
                try
                {
                    File.Delete(pdfFilePath);
                }
                catch
                {
                    // ignore cleanup errors
                }
            }
        }
    }

    // Returns a simple 1x1 PNG icon as a byte array
    private static byte[] GeneratePlaceholderIcon()
    {
        // Base64-encoded 1x1 transparent PNG
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        return Convert.FromBase64String(base64Png);
    }
}
