// Title: How to Embed Multiple Files as Attachments in a PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating placeholder files, adding each as an OLE object (icon) on a worksheet, setting the correct FileFormatType, enabling PdfSaveOptions.EmbedAttachments, and saving the workbook as a PDF that contains all embedded attachments.
// Keywords: Aspose.Cells PDF embed attachments | PdfSaveOptions EmbedAttachments C# | embed multiple OLE objects Aspose.Cells | export workbook with attached files | C# Aspose.Cells PDF attachment example | add OLE icons to Excel sheet | generate PDF package with source files
// Common Searches: embed multiple files in PDF using Aspose.Cells .NET | PdfSaveOptions EmbedAttachments example C# | add OLE objects to worksheet and export to PDF | Aspose.Cells attach Word Excel PDF to generated PDF | how to bundle source documents inside a PDF with Aspose
// Developer Intent: The developer wants to bundle several external documents (e.g., DOCX, XLSX, PDF) as embedded attachments inside a PDF generated from an Aspose.Cells workbook.
// Use Cases: Financial report PDF that includes supporting schedules, contracts, and audit trails as embedded files for reviewers. | Regulatory submission package where the PDF contains the original source documents for compliance verification. | Product documentation bundle that ships a specification PDF together with design drawings, spreadsheets, and reference manuals.
// AI Prompts: Show how to embed image or text files as PDF attachments using PdfSaveOptions.EmbedAttachments in Aspose.Cells. | Add error handling for missing or inaccessible files when creating OLE objects before PDF conversion. | Customize the size, position, and label of OLE icons while keeping the attachments embedded in the final PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating placeholder files, adding each as an OLE object (icon) on a worksheet, setting the correct FileFormatType, enabling PdfSaveOptions.EmbedAttachments, and saving the workbook as a PDF that contains all embedded attachments.
class EmbedMultipleAttachmentsToPdf
{
    static void Main()
    {
        // Paths of files to embed
        string[] filesToEmbed = new string[]
        {
            "sample1.docx",
            "sample2.xlsx",
            "sample3.pdf"
        };

        // Create simple placeholder files for the demo
        foreach (string path in filesToEmbed)
        {
            File.WriteAllText(path, $"Content of {Path.GetFileName(path)}");
        }

        // Create a new workbook and add a title cell
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF with multiple embedded attachments");

        // Add each external file as an OLE object (displayed as an icon)
        int startRow = 2;
        foreach (string filePath in filesToEmbed)
        {
            byte[] fileData = File.ReadAllBytes(filePath);
            int oleIndex = sheet.OleObjects.Add(startRow, 0, 200, 200, fileData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the correct file format based on extension
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            switch (ext)
            {
                case ".docx":
                    ole.FileFormatType = FileFormatType.Docx;
                    break;
                case ".xlsx":
                    ole.FileFormatType = FileFormatType.Xlsx;
                    break;
                case ".pdf":
                    ole.FileFormatType = FileFormatType.Pdf;
                    break;
                default:
                    ole.FileFormatType = FileFormatType.Unknown;
                    break;
            }

            ole.DisplayAsIcon = true;
            ole.Label = Path.GetFileName(filePath);

            startRow += 5; // leave space before the next icon
        }

        // Configure PDF save options to embed OLE attachments
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedAttachments = true;

        // Save the workbook as PDF with embedded attachments
        workbook.Save("MultipleAttachments.pdf", pdfOptions);

        // Clean up the temporary files created for the demo
        foreach (string path in filesToEmbed)
        {
            File.Delete(path);
        }
    }
}
