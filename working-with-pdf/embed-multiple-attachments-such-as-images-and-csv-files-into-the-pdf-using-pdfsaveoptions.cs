// Title: Embed Multiple Attachments (Image & CSV) into a PDF using Aspose.Cells PdfSaveOptions (C#)
// Description: Creates a workbook, adds an image and a CSV file as OLE objects, sets each object's format and icon display, enables the EmbedAttachments flag in PdfSaveOptions, and saves the workbook as a PDF that contains both files as embedded attachments. Temporary files are cleaned up after saving.
// Keywords: Aspose.Cells embed attachments PDF | PdfSaveOptions EmbedAttachments C# | add OLE objects worksheet | embed image in PDF Aspose.Cells | embed CSV in PDF Aspose.Cells | multiple file attachments PDF | C# Aspose.Cells PDF export
// Common Searches: how to embed multiple files in a PDF with Aspose.Cells | C# embed image and CSV as PDF attachments using PdfSaveOptions | Aspose.Cells PdfSaveOptions EmbedAttachments example | add OLE objects to Excel and export to PDF with attachments | save workbook as PDF with embedded files Aspose
// Developer Intent: The developer wants to bundle several files (e.g., an image and a CSV) inside a PDF generated from an Excel workbook using Aspose.Cells.
// Use Cases: Produce a financial report PDF that carries supporting charts (PNG) and raw data (CSV) for auditors. | Create an invoice PDF that includes product photos and a CSV of line‑item details for downstream processing. | Automate delivery of technical documentation by embedding reference diagrams and data files within a single PDF generated from a template.
// AI Prompts: Show a C# example that adds multiple OLE objects of different types to a worksheet and saves the workbook as a PDF with embedded attachments using Aspose.Cells. | Provide code to embed an image and a CSV file as PDF attachments via PdfSaveOptions in Aspose.Cells, including cleanup of temporary files. | Explain how the EmbedAttachments property works in PdfSaveOptions and how to configure OLE objects to display as icons.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds an image and a CSV file as OLE objects, sets each object's format and icon display, enables the EmbedAttachments flag in PdfSaveOptions, and saves the workbook as a PDF that contains both files as embedded attachments. Temporary files are cleaned up after saving.
class EmbedMultipleAttachmentsToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with Multiple Embedded Attachments");

            // Prepare a sample image file (minimal PNG header)
            string imageFile = "sampleImage.png";
            if (!File.Exists(imageFile))
            {
                File.WriteAllBytes(imageFile, new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
            }

            // Prepare a sample CSV file
            string csvFile = "sampleData.csv";
            if (!File.Exists(csvFile))
            {
                File.WriteAllText(csvFile, "Name,Age\nJohn,30\nAlice,25");
            }

            // Read file bytes for OLE embedding (API expects byte[])
            byte[] imageBytes = File.ReadAllBytes(imageFile);
            int imageOleIndex = worksheet.OleObjects.Add(5, 0, 200, 200, imageBytes);
            worksheet.OleObjects[imageOleIndex].FileFormatType = FileFormatType.Png;
            worksheet.OleObjects[imageOleIndex].DisplayAsIcon = true;

            byte[] csvBytes = File.ReadAllBytes(csvFile);
            int csvOleIndex = worksheet.OleObjects.Add(15, 0, 200, 200, csvBytes);
            worksheet.OleObjects[csvOleIndex].FileFormatType = FileFormatType.Csv;
            worksheet.OleObjects[csvOleIndex].DisplayAsIcon = true;

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with embedded attachments
            workbook.Save("MultipleAttachments.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Clean up temporary files
            try { if (File.Exists("sampleImage.png")) File.Delete("sampleImage.png"); } catch { }
            try { if (File.Exists("sampleData.csv")) File.Delete("sampleData.csv"); } catch { }
        }
    }
}
